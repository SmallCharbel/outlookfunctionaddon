// Required for Microsoft Graph client in Node.js environment
require('isomorphic-fetch');

const { Client } = require('@microsoft/microsoft-graph-client');

// Modify the message retrieval approach to use metadata search by recipient and subject
async function searchMessageByMetadata(client, subject, recipients) { // Removed contentSnippet and receivedTime from parameters
    const encodeOData = (str) => str ? str.replace(/'/g, "''") : ""; // Added check for null/undefined str

    // Build OData filter parts
    const filterParts = [];

    if (subject) { // Check if subject is provided
        filterParts.push(`subject eq '${encodeOData(subject)}'`);
    }

    // If recipients string is provided (semicolon-separated), use the first address for filtering
    if (recipients) {
        const firstRecipient = recipients.split(';')[0].trim();
        if (firstRecipient) {
            filterParts.push(`toRecipients/any(r: r/emailAddress/address eq '${encodeOData(firstRecipient)}')`);
        }
    }

    // If no filter parts are available (e.g., neither subject nor recipient provided),
    // it's probably not a good idea to search all messages.
    // However, the original logic proceeded if subject and receivedTime were present.
    // We'll proceed if at least one (subject or recipient) is present.
    if (filterParts.length === 0) {
        // Optionally, handle this case by returning null or throwing an error,
        // as searching without filters can be very broad.
        // For now, let's assume the calling logic ensures at least one is usually present.
        context.log.warn("SearchMessageByMetadata called without subject or recipients. This may lead to a broad search or errors.");
        // return null; // Or throw new Error("Subject or recipient is required for search.");
    }

    const filter = filterParts.join(' and ');

    try {
        const apiCall = client.api('/me/messages');
        if (filter) { // Only apply filter if it's not empty
            apiCall.filter(filter);
        }

        // Removed .orderby('receivedDateTime desc') to reduce complexity
        // Kept .top(1) to get the most relevant single message
        // If no specific ordering, top(1) will return an arbitrary message matching the filter.
        // If a specific message is needed, some form of ordering or more specific criteria might be required.
        const response = await apiCall
            .top(1)
            .select('id')
            .get();

        if (response.value && response.value.length > 0) {
            return response.value[0].id;
        }
        return null;
    } catch (err) {
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
        // The condition for search now relies on subject OR recipients being present.
        // The original logic checked `subject && receivedTime`.
        // We adjust this to reflect the simplified search parameters.
        if ((!messageId || messageId === '') && useMetadataSearch && (subject || recipients)) {
            context.log(`Searching for email with subject: ${subject}, recipients: ${recipients}`); // Removed receivedTime from log
            try {
                // Pass only client, subject, and recipients to the modified function
                messageId = await searchMessageByMetadata(client, subject, recipients);
                if (messageId) {
                    context.log(`Found message with ID: ${messageId}`);
                } else {
                    context.log.error(`No messages found matching metadata criteria`);
                    context.res = {
                        status: 404,
                        body: {
                            success: false,
                            error: `No messages found with subject: ${subject}, recipients: ${recipients}`
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
            context.log.warn("No message ID provided or found after metadata search attempt."); // Changed log level
            context.res = {
                status: 400, // Or 404 if search was attempted but failed to find.
                body: {
                    error: "Message ID is required and could not be determined via metadata search.",
                }
            };
            return;
        }

        // Validate the message ID format
        if (!validateMessageId(messageId)) {
            context.log.error(`Invalid message ID format: ${messageId}`);

            // Try to translate the ID if it seems to be in EWS format
            if (messageId.includes('/')) {
                try {
                    context.log("Attempting to convert Exchange ID format to REST format");
                    messageId = await convertExchangeId(client, messageId);
                    context.log(`Converted message ID: ${messageId}`);
                } catch (error) {
                    context.log.error(`Failed to convert Exchange ID: ${error.message}`);
                    // Continue with the original ID as a fallback
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

                // Log attachment details for debugging
                attachments.forEach((attachment, index) => {
                    context.log(`Attachment ${index + 1}: Name=${attachment.name}, Type=${attachment["@odata.type"]}, Size=${attachment.size || 'unknown'}`);
                });
            }

            // Create a new message draft
            context.log("Creating new message draft...");
            const newMessage = {
                subject: `${message.subject}`, // Keep the original subject without adding "FW:"
                body: {
                    contentType: message.body.contentType,
                    content: message.body.content
                },
                // Preserve all original recipients
                toRecipients: message.toRecipients || [],
                // Include original CC recipients
                ccRecipients: message.ccRecipients || [],
                // Copy importance from original
                importance: message.importance || "normal"
            };

            // Create the draft message
            const draftMessage = await client.api('/me/messages').post(newMessage);

            // Add attachments if any
            if (attachments.length > 0) {
                context.log(`Adding ${attachments.length} attachments...`);

                for (const attachment of attachments) {
                    context.log(`Adding attachment: ${attachment.name}`);

                    try {
                        // Create proper attachment object based on type
                        const attachmentData = {
                            "@odata.type": attachment["@odata.type"],
                            name: attachment.name,
                            contentType: attachment.contentType
                        };

                        // Add contentBytes for file attachments
                        if (attachment["@odata.type"] === "#microsoft.graph.fileAttachment") {
                            attachmentData.contentBytes = attachment.contentBytes;
                        }
                        // For item attachments
                        else if (attachment["@odata.type"] === "#microsoft.graph.itemAttachment") {
                            context.log(`Item attachment detected: ${attachment.name}`);
                            attachmentData.item = attachment.item; // Ensure 'item' property exists and is structured correctly
                        }
                        // For reference attachments
                        else if (attachment["@odata.type"] === "#microsoft.graph.referenceAttachment") {
                            attachmentData.providerType = attachment.providerType; // Ensure correct property name (providerType vs referenceAttachmentType)
                            attachmentData.sourceUrl = attachment.sourceUrl;
                            // attachmentData.permission = attachment.permission; // Check if these are always present
                            // attachmentData.isFolder = attachment.isFolder;
                        }


                        await client.api(`/me/messages/${draftMessage.id}/attachments`).post(attachmentData);
                        context.log(`Successfully added attachment: ${attachment.name}`);
                    } catch (attachError) {
                        context.log.error(`Error adding attachment ${attachment.name}: ${attachError.message} - ${JSON.stringify(attachError.body)}`);
                        // Continue with other attachments even if one fails
                    }
                }
            }

            // Send the message
            context.log("Sending the new message...");
            await client.api(`/me/messages/${draftMessage.id}/send`).post({});

            // Move original message to deleted items
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
             // Check if error is a GraphError and has a more specific code
            let errorMessage = `Message not found or error accessing: ${error.message}`;
            if (error.statusCode && error.code) {
                errorMessage = `Graph API Error: ${error.code} - ${error.message}`;
            }

            context.res = {
                status: (error.statusCode === 404 ? 404 : 500), // More specific status if 404
                body: {
                    success: false, // Ensure success is false on error
                    error: errorMessage,
                    messageIdUsed: messageId
                }
            };
        }
    } catch (error) {
        context.log.error(`Error: ${error.message}`);
        context.res = {
            status: 500,
            body: { // Ensure body is an object for consistency
                success: false,
                error: `Error forwarding email: ${error.message}`
            }
        };
    }
};

// Helper function to create authenticated client
function getAuthenticatedClient(accessToken) {
    // const { Client } is already defined at the top
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
    // Removed length check and numeric check as REST IDs can vary.
    // The primary check is for EWS-like IDs containing '/'.
    // Graph API will ultimately validate the ID.
    if (id.includes('/')) return false; // This indicates it might be an EWS ID
    return true;
}

// Convert Exchange ID to REST format using translateExchangeIds
async function convertExchangeId(client, exchangeId) {
    try {
        const response = await client.api('/me/translateExchangeIds').post({
            inputIds: [exchangeId],
            targetIdType: "restId",
            sourceIdType: "ewsId" // Assuming the ID is ewsId if it contains '/'
        });
        if (response?.value?.length > 0 && response.value[0].targetId) {
            return response.value[0].targetId;
        }
        throw new Error("No translated ID returned or targetId is missing.");
    } catch (error) {
        // Log the full error if possible
        const errorMessage = error.body ? JSON.stringify(error.body) : error.message;
        throw new Error(`Translation failed: ${errorMessage}`);
    }
}
