// Required for Microsoft Graph client in Node.js environment
require('isomorphic-fetch');

const { Client } = require('@microsoft/microsoft-graph-client');

// Modify the message retrieval approach to use metadata search by recipient, subject, and receivedTime
async function searchMessageByMetadata(client, subject, recipients, receivedTime, context) { // Added receivedTime parameter
    const encodeOData = (str) => str ? str.replace(/'/g, "''") : "";

    const initialFilterParts = [];
    if (subject) {
        initialFilterParts.push(`subject eq '${encodeOData(subject)}'`);
    }

    // Add receivedTime to the initial filter if provided
    if (receivedTime) {
        // Ensure receivedTime is a valid ISO string, Graph API expects this format.
        // Example: "2024-06-05T12:30:00Z"
        initialFilterParts.push(`receivedDateTime eq ${encodeOData(receivedTime)}`);
    }

    const fullRecipientList = recipients ? recipients.split(';').map(r => r.trim().toLowerCase()).filter(r => r) : [];

    if (fullRecipientList.length > 0) {
        // Use only the first recipient for the initial OData filter to reduce complexity further if needed
        // However, if receivedTime is very specific, this might be okay.
        // For now, keeping the first recipient logic for robustness.
        initialFilterParts.push(`toRecipients/any(r: r/emailAddress/address eq '${encodeOData(fullRecipientList[0])}')`);
    }


    if (initialFilterParts.length === 0) {
        if (context && context.log) {
            context.log.warn("SearchMessageByMetadata called without any criteria for initial filter.");
        }
        return null; // Avoid searching without any initial filters
    }
    // If only receivedTime and one recipient are provided, but no subject, that's a valid filter.
    // If only subject and receivedTime, that's also valid.

    const initialFilter = initialFilterParts.join(' and ');
    if (context && context.log) {
        context.log.info(`Constructed initial OData filter: ${initialFilter}`);
    }

    try {
        // Fetch top N messages (e.g., 5) matching the simplified initial filter
        // We need to select toRecipients to perform the second stage of filtering
        const response = await client.api('/me/messages')
            .filter(initialFilter)
            .orderby('receivedDateTime desc') // Order by is kept; if time is exact, this might be redundant but harmless
            .top(5) // Fetch a small batch for secondary filtering
            .select('id,receivedDateTime,subject,toRecipients')
            .get();

        if (response.value && response.value.length > 0) {
            if (context && context.log) {
                context.log.info(`Initial query returned ${response.value.length} messages. Performing secondary filtering for all recipients if necessary...`);
            }

            // Secondary filtering: iterate through the results and check for all recipients
            // This is mainly relevant if fullRecipientList has more than one recipient,
            // or if the initial filter didn't use any recipient filter.
            for (const message of response.value) {
                // If fullRecipientList is empty or has only one (which was in initial filter),
                // this secondary check for recipients is less critical but kept for consistency.
                if (fullRecipientList.length > 1) { // Only perform full recipient check if there were multiple expected recipients
                    if (!message.toRecipients || message.toRecipients.length === 0) {
                         if (context && context.log) { context.log.info(`Message ${message.id} skipped in secondary filter: no recipients found in message, but expected ${fullRecipientList.length}.`);}
                        continue; // If we expect multiple recipients but message has none, skip
                    }

                    const messageRecipientsSet = new Set(
                        (message.toRecipients || []).map(r => r.emailAddress && r.emailAddress.address ? r.emailAddress.address.toLowerCase() : '')
                    );

                    let allRecipientsFound = true;
                    for (const expectedRecipient of fullRecipientList) {
                        if (!messageRecipientsSet.has(expectedRecipient)) {
                            allRecipientsFound = false;
                             if (context && context.log) { context.log.info(`Message ${message.id} failed secondary recipient check: missing ${expectedRecipient}. Message recipients: ${Array.from(messageRecipientsSet).join(', ')}`);}
                            break;
                        }
                    }
                    if (!allRecipientsFound) {
                        continue; // Skip this message if not all expected recipients are found
                    }
                }


                // If we reach here, the message matches the initial filter (including time and first recipient if applicable)
                // AND it has passed the secondary check for all recipients (if there were multiple).
                if (context && context.log) {
                    context.log.info(`Found message with ID: ${message.id} (Subject: ${message.subject}, Received: ${message.receivedDateTime}) matching all criteria.`);
                }
                return message.id; // This is the latest message matching all criteria
            }
            if (context && context.log) {
                context.log.info("No messages matched all recipient criteria after secondary filtering of top results (or primary if only one recipient).");
            }
        } else {
            if (context && context.log) {
                context.log.info("Initial OData query returned no messages.");
            }
        }
        return null; // No message found matching all criteria
    } catch (err) {
        if (context && context.log) {
            context.log.error(`Error during message search stages: ${err.message}. Initial filter used: ${initialFilter}`);
        }
        throw new Error(`Error searching by metadata: ${err.message}`);
    }
}

// Main function handler
module.exports = async function (context, req) {
    context.log("Processing email forwarding request");

    try {
        context.log(`Headers received: ${JSON.stringify(req.headers)}`);
        context.log(`Request body: ${JSON.stringify(req.body || {})}`);

        const authHeader = req.headers.authorization || '';
        if (!authHeader.startsWith('Bearer ')) {
            context.log.error("No authorization token provided");
            context.res = { status: 401, body: "Unauthorized: No token provided" };
            return;
        }

        const accessToken = authHeader.substring(7);
        const client = getAuthenticatedClient(accessToken);
        context.log("Creating Graph client with delegated token");

        const {
            messageId: providedId,
            subject,
            recipients,
            contentSnippet,
            receivedTime, // This is the receivedTime from the request
            userEmail,
            useMetadataSearch
        } = req.body || {};

        let messageId = providedId;
        context.log(`Original message ID from request: ${messageId || 'Not provided'}`);

        if ((!messageId || messageId === '') && useMetadataSearch && (subject || recipients || receivedTime)) { // Added receivedTime to condition
            context.log(`Searching for email with subject: "${subject}", recipients: "${recipients}", receivedTime: "${receivedTime}"`);
            try {
                // Pass receivedTime to searchMessageByMetadata
                messageId = await searchMessageByMetadata(client, subject, recipients, receivedTime, context);
                if (messageId) {
                    context.log(`Found message with ID: ${messageId}`);
                } else {
                    context.log.error(`No messages found matching metadata criteria: Subject="${subject}", Recipients="${recipients}", ReceivedTime="${receivedTime}"`);
                    context.res = {
                        status: 404,
                        body: {
                            success: false,
                            error: `No messages found with subject: "${subject}", recipients: "${recipients}", and receivedTime: "${receivedTime}"`
                        }
                    };
                    return;
                }
            } catch (error) {
                context.log.error(`Error searching for message: ${error.message}`);
                context.res = {
                    status: 500,
                    body: { success: false, error: `Error searching for message: ${error.message}` }
                };
                return;
            }
        }

        if (!messageId) {
            context.log.warn("No message ID provided or found after metadata search attempt.");
            context.res = {
                status: 400,
                body: { success: false, error: "Message ID is required and could not be determined via metadata search." }
            };
            return;
        }

        if (!validateMessageId(messageId)) {
            context.log.error(`Invalid message ID format: ${messageId}`);
            if (messageId.includes('/')) {
                try {
                    context.log("Attempting to convert Exchange ID format to REST format");
                    messageId = await convertExchangeId(client, messageId, context);
                    context.log(`Converted message ID: ${messageId}`);
                } catch (error) {
                    context.log.error(`Failed to convert Exchange ID: ${error.message}`);
                    // Fallback to original ID might still fail, but we try.
                }
            }
        }

        context.log(`Fetching original message with ID: ${messageId}`);
        try {
            const message = await client.api(`/me/messages/${messageId}`)
                .select('subject,body,toRecipients,ccRecipients,bccRecipients,from,hasAttachments,importance,isRead')
                .get();
            context.log(`Successfully retrieved message: ${message.subject}`);

            let attachments = [];
            if (message.hasAttachments) {
                context.log("Fetching attachments...");
                const attachmentsResponse = await client.api(`/me/messages/${messageId}/attachments`).get();
                attachments = attachmentsResponse.value;
                context.log(`Found ${attachments.length} attachments`);
                attachments.forEach((att, idx) => context.log(`Attachment ${idx + 1}: Name=${att.name}, Type=${att["@odata.type"]}`));
            }

            context.log("Creating new message draft...");
            const newMessage = {
                subject: `${message.subject}`,
                body: { contentType: message.body.contentType, content: message.body.content },
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
            await client.api(`/me/messages/${messageId}/move`).post({ destinationId: "deleteditems" });

            context.log("Process completed successfully");
            context.res = {
                status: 200,
                body: { success: true, message: "Email forwarded successfully" }
            };
        } catch (error) {
            context.log.error(`Error accessing message: ${error.message} - Message ID used: ${messageId}`);
            let errMsg = `Error accessing message: ${error.message}`;
            if (error.statusCode && error.code) errMsg = `Graph API Error: ${error.code} - ${error.message}`;
            context.res = {
                status: (error.statusCode === 404 ? 404 : 500),
                body: { success: false, error: errMsg, messageIdUsed: messageId }
            };
        }
    } catch (error) {
        context.log.error(`Error: ${error.message}`);
        context.res = {
            status: 500,
            body: { success: false, error: `Error forwarding email: ${error.message}` }
        };
    }
};

function getAuthenticatedClient(accessToken) {
    const client = Client.init({ authProvider: (done) => done(null, accessToken) });
    return client;
}

function validateMessageId(id) {
    if (!id) return false;
    if (id.includes('/')) return false;
    return true;
}

async function convertExchangeId(client, exchangeId, context) {
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
        const errMsg = error.body ? JSON.stringify(error.body) : error.message;
        if (context && context.log) context.log.error(`Translation failed for ID ${exchangeId}: ${errMsg}`);
        throw new Error(`Translation failed: ${errMsg}`);
    }
}
