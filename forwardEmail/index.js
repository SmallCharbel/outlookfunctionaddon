// Required for Microsoft Graph client in Node.js environment
require('isomorphic-fetch');

const { Client } = require('@microsoft/microsoft-graph-client');

// Modify the message retrieval approach to use metadata search by recipient and subject
async function searchMessageByMetadata(client, subject, recipients, context) { // Added context for logging
    const encodeOData = (str) => str ? str.replace(/'/g, "''") : "";

    // Build OData filter parts
    const filterParts = [];

    if (subject) {
        filterParts.push(`subject eq '${encodeOData(subject)}'`);
    }

    // If recipients string is provided (semicolon-separated), add a filter for each recipient
    if (recipients) {
        const recipientList = recipients.split(';').map(r => r.trim()).filter(r => r); // Get a clean list of recipients
        if (recipientList.length > 0) {
            recipientList.forEach(recipientAddress => {
                // Add a filter condition for each recipient
                // This ensures the message was sent to ALL specified recipients.
                // Note: The Graph API's $filter on `toRecipients` checks if *any* of the emailAddress objects in the collection match.
                // To ensure a message was sent to *all* specified recipients, each must be an 'and' condition.
                filterParts.push(`toRecipients/any(r: r/emailAddress/address eq '${encodeOData(recipientAddress)}')`);
            });
        }
    }

    if (filterParts.length === 0) {
        if (context && context.log) { // Check if context and context.log are available
            context.log.warn("SearchMessageByMetadata called without subject or recipients. This may lead to a broad search or errors.");
        }
        return null; // Avoid searching without any filters
    }

    const filter = filterParts.join(' and ');
    if (context && context.log) {
        context.log.info(`Constructed OData filter: ${filter}`);
    }


    try {
        const apiCall = client.api('/me/messages');
        if (filter) {
            apiCall.filter(filter);
        }

        // Add orderby to get the latest message first
        // Keep .top(1) to get the most relevant single (latest) message
        const response = await apiCall
            .orderby('receivedDateTime desc') // Sort by received time, latest first
            .top(1)
            .select('id,receivedDateTime,subject') // Added receivedDateTime and subject for logging/verification if needed
            .get();

        if (response.value && response.value.length > 0) {
            if (context && context.log) {
                 context.log.info(`Found message with ID: ${response.value[0].id}, Subject: ${response.value[0].subject}, Received: ${response.value[0].receivedDateTime}`);
            }
            return response.value[0].id;
        }
        if (context && context.log) {
            context.log.info("No messages found matching the specified criteria.");
        }
        return null;
    } catch (err) {
         if (context && context.log) {
            context.log.error(`Error searching by metadata: ${err.message}. Filter used: ${filter}`);
        }
        throw new Error(`Error searching by metadata: ${err.message}`);
    }
}

// Main function handler
module.exports = async function (context, req) {
    context.log("Processing email forwarding request");

    try {
        // Log headers for debugging
        context.log(`Headers received: ${JSON.stringify(req.headers)}`);

        // Log request body for debugging
        context.log(`Request body: ${JSON.stringify(req.body || {})}`);

        // Get authorization token from header
        const authHeader = req.headers.authorization || '';
        if (!authHeader.startsWith('Bearer ')) {
            context.log.error("No authorization token provided");
            context.res = {
                status: 401,
                body: "Unauthorized: No token provided"
            };
            return;
        }

        context.log("Found authorization header");
        const accessToken = authHeader.substring(7); // Remove 'Bearer ' prefix

        // Create Microsoft Graph client
        const client = getAuthenticatedClient(accessToken);
        context.log("Creating Graph client with delegated token");

        // Extract payload fields
        const {
            messageId: providedId,
            subject,
            recipients,
            contentSnippet, // Still received, but not used in the simplified search
            receivedTime,   // Still received, but not used in the simplified search
            userEmail,
            useMetadataSearch
        } = req.body || {};

        let messageId = providedId;
        context.log(`Original message ID from request: ${messageId || 'Not provided'}`);

        // If metadata search is requested, use it to find the message ID
        if ((!messageId || messageId === '') && useMetadataSearch && (subject || recipients)) {
            context.log(`Searching for email with subject: "${subject}", recipients: "${recipients}"`);
            try {
                // Pass context to searchMessageByMetadata for logging
                messageId = await searchMessageByMetadata(client, subject, recipients, context);
                if (messageId) {
                    context.log(`Found message with ID: ${messageId}`);
                } else {
                    context.log.error(`No messages found matching metadata criteria: Subject="${subject}", Recipients="${recipients}"`);
                    context.res = {
                        status: 404,
                        body: {
                            success: false,
                            error: `No messages found with subject: "${subject}", and all recipients: "${recipients}"`
                        }
                    };
                    return;
                }
            } catch (error) {
                context.log.error(`Error searching for message: ${error.message}`);
                context.res = {
                    status: 500,
                    body: {
                        success: false,
                        error: `Error searching for message: ${error.message}`
                    }
                };
                return;
            }
        }

        // If still no message ID provided after metadata search, return error
        if (!messageId) {
            context.log.warn("No message ID provided or found after metadata search attempt.");
            context.res = {
                status: 400,
                body: {
                    success: false, // ensure success is false
                    error: "Message ID is required and could not be determined via metadata search.",
                }
            };
            return;
        }

        // Validate the message ID format
        if (!validateMessageId(messageId)) {
            context.log.error(`Invalid message ID format: ${messageId}`);

            if (messageId.includes('/')) {
                try {
                    context.log("Attempting to convert Exchange ID format to REST format");
                    messageId = await convertExchangeId(client, messageId, context); // Pass context
                    context.log(`Converted message ID: ${messageId}`);
                } catch (error) {
                    context.log.error(`Failed to convert Exchange ID: ${error.message}`);
                }
            }
        }

        // Fetch the original message
        context.log(`Fetching original message with ID: ${messageId}`);
        try {
            const message = await client.api(`/me/messages/${messageId}`)
                .select('subject,body,toRecipients,ccRecipients,bccRecipients,from,hasAttachments,importance,isRead')
                .get();

            context.log(`Successfully retrieved message: ${message.subject}`);

            // Fetch attachments if any
            let attachments = [];
            if (message.hasAttachments) {
                context.log("Fetching attachments...");
                const attachmentsResponse = await client.api(`/me/messages/${messageId}/attachments`).get();
                attachments = attachmentsResponse.value;
                context.log(`Found ${attachments.length} attachments`);

                attachments.forEach((attachment, index) => {
                    context.log(`Attachment ${index + 1}: Name=${attachment.name}, Type=${attachment["@odata.type"]}, Size=${attachment.size || 'unknown'}`);
                });
            }

            // Create a new message draft
            context.log("Creating new message draft...");
            const newMessage = {
                subject: `${message.subject}`,
                body: {
                    contentType: message.body.contentType,
                    content: message.body.content
                },
                toRecipients: message.toRecipients || [],
                ccRecipients: message.ccRecipients || [],
                importance: message.importance || "normal"
            };

            const draftMessage = await client.api('/me/messages').post(newMessage);

            if (attachments.length > 0) {
                context.log(`Adding ${attachments.length} attachments...`);
                for (const attachment of attachments) {
                    context.log(`Adding attachment: ${attachment.name}`);
                    try {
                        const attachmentData = {
                            "@odata.type": attachment["@odata.type"],
                            name: attachment.name,
                            contentType: attachment.contentType
                        };

                        if (attachment["@odata.type"] === "#microsoft.graph.fileAttachment") {
                            attachmentData.contentBytes = attachment.contentBytes;
                        } else if (attachment["@odata.type"] === "#microsoft.graph.itemAttachment") {
                            context.log(`Item attachment detected: ${attachment.name}`);
                            attachmentData.item = attachment.item;
                        } else if (attachment["@odata.type"] === "#microsoft.graph.referenceAttachment") {
                            attachmentData.providerType = attachment.providerType;
                            attachmentData.sourceUrl = attachment.sourceUrl;
                        }
                        await client.api(`/me/messages/${draftMessage.id}/attachments`).post(attachmentData);
                        context.log(`Successfully added attachment: ${attachment.name}`);
                    } catch (attachError) {
                        context.log.error(`Error adding attachment ${attachment.name}: ${attachError.message} - ${JSON.stringify(attachError.body)}`);
                    }
                }
            }

            context.log("Sending the new message...");
            await client.api(`/me/messages/${draftMessage.id}/send`).post({});

            context.log("Moving original message to deleted items...");
            await client.api(`/me/messages/${messageId}/move`).post({
                destinationId: "deleteditems"
            });

            context.log("Process completed successfully");
            context.res = {
                status: 200,
                body: {
                    success: true,
                    message: "Email forwarded successfully"
                }
            };
        } catch (error) {
            context.log.error(`Error accessing message: ${error.message} - Message ID used: ${messageId}`);
            let errorMessage = `Message not found or error accessing: ${error.message}`;
            if (error.statusCode && error.code) {
                errorMessage = `Graph API Error: ${error.code} - ${error.message}`;
            }
            context.res = {
                status: (error.statusCode === 404 ? 404 : 500),
                body: {
                    success: false,
                    error: errorMessage,
                    messageIdUsed: messageId
                }
            };
        }
    } catch (error) {
        context.log.error(`Error: ${error.message}`);
        context.res = {
            status: 500,
            body: {
                success: false,
                error: `Error forwarding email: ${error.message}`
            }
        };
    }
};

// Helper function to create authenticated client
function getAuthenticatedClient(accessToken) {
    const client = Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        }
    });
    return client;
}

// Validate message ID format
function validateMessageId(id) {
    if (!id) return false;
    if (id.includes('/')) return false;
    return true;
}

// Convert Exchange ID to REST format using translateExchangeIds
async function convertExchangeId(client, exchangeId, context) { // Added context for logging
    try {
        const response = await client.api('/me/translateExchangeIds').post({
            inputIds: [exchangeId],
            targetIdType: "restId",
            sourceIdType: "ewsId"
        });
        if (response?.value?.length > 0 && response.value[0].targetId) {
            return response.value[0].targetId;
        }
        throw new Error("No translated ID returned or targetId is missing.");
    } catch (error) {
        const errorMessage = error.body ? JSON.stringify(error.body) : error.message;
         if (context && context.log) { // Check if context and context.log are available
            context.log.error(`Translation failed for ID ${exchangeId}: ${errorMessage}`);
        }
        throw new Error(`Translation failed: ${errorMessage}`);
    }
}
