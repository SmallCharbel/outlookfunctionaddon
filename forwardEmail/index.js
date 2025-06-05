// Required for Microsoft Graph client in Node.js environment
require('isomorphic-fetch');

const { Client } = require('@microsoft/microsoft-graph-client');

// Modify the message retrieval approach to use metadata search, primarily by receivedTime
async function searchMessageByMetadata(client, subjectFromRequest, recipientsFromRequest, receivedTimeFromRequest, context) {
    const encodeOData = (str) => str ? str.replace(/'/g, "''") : "";

    // Validate receivedTimeFromRequest: It's critical and must be a non-empty string.
    if (!receivedTimeFromRequest || typeof receivedTimeFromRequest !== 'string' || receivedTimeFromRequest.trim() === '') {
        if (context && context.log) {
            context.log.error("searchMessageByMetadata: receivedTimeFromRequest is missing, not a string, or empty. It is mandatory and must be a valid ISO date string for a precise search.");
        }
        // Throw an error to ensure this problematic case doesn't proceed.
        throw new Error("receivedTimeFromRequest must be a valid ISO date string and is required for metadata search to uniquely identify the message.");
    }

    // PRIMARY ODATA FILTER: Filter *only* by the exact receivedDateTime.
    // This assumes receivedTimeFromRequest is a valid ISO string like "2024-06-05T12:30:00Z".
    // encodeOData is used here defensively for the value, though for a clean ISO string it should have no effect.
    const initialODataFilter = `receivedDateTime eq ${encodeOData(receivedTimeFromRequest)}`;

    if (context && context.log) {
        context.log.info(`Constructed OData filter for Graph API: ${initialODataFilter}`);
    }

    try {
        // Fetch a small number of messages matching the exact timestamp.
        const response = await client.api('/me/messages')
            .filter(initialODataFilter)
            .select('id,receivedDateTime,subject,toRecipients')
            .top(5) // Fetch up to 5 as a safety net for rare exact timestamp collisions or slight precision variations.
            .get();

        if (response.value && response.value.length > 0) {
            if (context && context.log) {
                context.log.info(`Initial query by exact receivedTime ("${receivedTimeFromRequest}") returned ${response.value.length} message(s). Performing detailed validation...`);
            }

            const fullRecipientList = recipientsFromRequest ? recipientsFromRequest.split(';').map(r => r.trim().toLowerCase()).filter(r => r) : [];

            for (const message of response.value) {
                // Sanity check: Ensure the receivedDateTime from the message matches, accounting for potential minor string differences if any.
                // This is mostly for debugging; the primary filter should ensure this.
                if (message.receivedDateTime !== receivedTimeFromRequest) {
                    if (context && context.log) {
                        context.log.warn(`Message ${message.id}: Timestamp slight mismatch after query. Expected: "${receivedTimeFromRequest}", Actual: "${message.receivedDateTime}". Proceeding with validation if subjects/recipients match.`);
                        // Depending on strictness, you might 'continue' here if an exact string match is required.
                        // However, Graph API's `eq` on DateTimeOffset should handle precision correctly.
                    }
                }

                // 1. Validate Subject (if subjectFromRequest is provided)
                let subjectMatch = !subjectFromRequest; // True if no subject to check.
                if (subjectFromRequest) {
                    const messageSubjectNormalized = message.subject ? message.subject.trim().toLowerCase() : "";
                    const requestSubjectNormalized = subjectFromRequest.trim().toLowerCase();
                    if (messageSubjectNormalized === requestSubjectNormalized) {
                        subjectMatch = true;
                    } else {
                        if (context && context.log) {
                            context.log.info(`Message ${message.id} (Received: ${message.receivedDateTime}): Subject mismatch. Expected: "${requestSubjectNormalized}", Actual: "${messageSubjectNormalized}".`);
                        }
                    }
                }
                if (!subjectMatch) continue;

                // 2. Validate All Recipients (if recipientsFromRequest is provided)
                let recipientsMatch = fullRecipientList.length === 0; // True if no recipients to check.
                if (fullRecipientList.length > 0) { // Only validate if recipients were expected.
                    if (message.toRecipients && message.toRecipients.length > 0) {
                        const messageRecipientsSet = new Set(
                            message.toRecipients.map(r => r.emailAddress && r.emailAddress.address ? r.emailAddress.address.toLowerCase() : null).filter(Boolean)
                        );
                        
                        let allExpectedRecipientsFound = true;
                        for (const expectedRecipient of fullRecipientList) {
                            if (!messageRecipientsSet.has(expectedRecipient)) {
                                allExpectedRecipientsFound = false;
                                if (context && context.log) {
                                    context.log.info(`Message ${message.id} (Received: ${message.receivedDateTime}): Recipient mismatch. Expected recipient "${expectedRecipient}" not found in message recipients: [${Array.from(messageRecipientsSet).join(', ')}].`);
                                }
                                break;
                            }
                        }
                        if (allExpectedRecipientsFound) {
                            recipientsMatch = true;
                        }
                    } else {
                        // Message has no recipients, but we expected some.
                        if (context && context.log) context.log.info(`Message ${message.id} (Received: ${message.receivedDateTime}): Recipient mismatch. Expected ${fullRecipientList.length} recipients, but message has none.`);
                    }
                }
                if (!recipientsMatch) continue;


                if (context && context.log) {
                    context.log.info(`SUCCESS: Message ${message.id} (Received: "${message.receivedDateTime}", Subject: "${message.subject}") passed all client-side validations.`);
                }
                return message.id;
            }

            if (context && context.log) {
                context.log.info("No messages passed the detailed client-side validation (subject/recipients) after initial query by receivedTime.");
            }
        } else {
            if (context && context.log) {
                context.log.info(`Initial OData query by exact receivedTime ("${receivedTimeFromRequest}") returned no messages.`);
            }
        }
        return null;

    } catch (err) {
        if (context && context.log) {
            context.log.error(`Error during Graph API call in searchMessageByMetadata (Filter was: "${initialODataFilter}"): ${err.message}`);
        }
        throw new Error(`Error searching message by metadata: ${err.message}`);
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
            receivedTime, 
            userEmail,
            useMetadataSearch
        } = req.body || {};

        let messageId = providedId;
        context.log(`Original message ID from request: ${messageId || 'Not provided'}`);

        if ((!messageId || messageId === '') && useMetadataSearch) {
            context.log(`Attempting metadata search for email with Subject: "${subject}", Recipients: "${recipients}", ReceivedTime: "${receivedTime}"`);
            try {
                // searchMessageByMetadata now throws if receivedTime is invalid, so it will be caught here.
                messageId = await searchMessageByMetadata(client, subject, recipients, receivedTime, context);
                if (messageId) {
                    context.log(`Metadata search successful. Found message with ID: ${messageId}`);
                } else {
                    // This path is less likely if searchMessageByMetadata throws on invalid time or returns null on no match.
                    context.log.error(`Metadata search did not find a matching message. Criteria: Subject="${subject}", Recipients="${recipients}", ReceivedTime="${receivedTime}"`);
                    context.res = {
                        status: 404,
                        body: {
                            success: false,
                            error: `No message found matching all provided criteria (Subject: "${subject}", Recipients: "${recipients}", ReceivedTime: "${receivedTime}")`
                        }
                    };
                    return;
                }
            } catch (error) { // Catch errors from searchMessageByMetadata (e.g., invalid receivedTime or Graph API errors)
                context.log.error(`Error during metadata search: ${error.message}`);
                context.res = {
                    status: error.message.includes("receivedTimeFromRequest must be a valid ISO date string") ? 400 : 500,
                    body: { success: false, error: `Error searching for message: ${error.message}` }
                };
                return;
            }
        }

        if (!messageId) {
            context.log.warn("No message ID available (either not provided or metadata search failed/not used).");
            context.res = {
                status: 400,
                body: { success: false, error: "Message ID is required and could not be determined." }
            };
            return;
        }

        // Validate messageId format (heuristic) and attempt EWS ID conversion if needed
        if (!validateMessageId(messageId)) { // validateMessageId checks if it might be EWS
            context.log.warn(`Potentially invalid or EWS message ID format detected: ${messageId}.`);
            if (messageId.includes('/')) { 
                try {
                    context.log("Attempting to convert Exchange ID (EWS) format to REST format");
                    messageId = await convertExchangeId(client, messageId, context);
                    context.log(`Successfully converted message ID to REST format: ${messageId}`);
                } catch (error) {
                    context.log.error(`Failed to convert EWS message ID "${messageId}" to REST format: ${error.message}. Proceeding with original ID, which may fail.`);
                }
            } else {
                 context.log.error(`Message ID format "${messageId}" is not recognized as EWS and failed basic validation. It may not work with Graph API.`);
            }
        }

        context.log(`Fetching original message with resolved ID: ${messageId}`);
        try {
            const message = await client.api(`/me/messages/${messageId}`)
                .select('subject,body,toRecipients,ccRecipients,bccRecipients,from,hasAttachments,importance,isRead')
                .get();
            context.log(`Successfully retrieved original message: "${message.subject}" (ID: ${message.id})`);

            let attachments = [];
            if (message.hasAttachments) {
                context.log("Original message has attachments. Fetching attachment details...");
                const attachmentsResponse = await client.api(`/me/messages/${messageId}/attachments`).get();
                attachments = attachmentsResponse.value || [];
                context.log(`Found ${attachments.length} attachments in the original message.`);
                attachments.forEach((att, idx) => context.log(`Attachment ${idx + 1}: Name="${att.name}", Type="${att["@odata.type"]}", Size=${att.size || 'N/A'}`));
            }

            context.log("Creating new message draft for forwarding...");
            const newMessage = {
                subject: `${message.subject}`,
                body: { contentType: message.body.contentType, content: message.body.content },
                toRecipients: message.toRecipients || [],
                ccRecipients: message.ccRecipients || [],
                importance: message.importance || "normal"
            };
            const draftMessage = await client.api('/me/messages').post(newMessage);
            context.log(`New draft message created with ID: ${draftMessage.id}. Subject: "${draftMessage.subject}"`);

            if (attachments.length > 0) {
                context.log(`Adding ${attachments.length} attachments to the new draft...`);
                for (const attachment of attachments) {
                    context.log(`Processing attachment: "${attachment.name}" (Type: ${attachment["@odata.type"]})`);
                    try {
                        const attachmentData = {
                            "@odata.type": attachment["@odata.type"],
                            name: attachment.name,
                            contentType: attachment.contentType,
                        };

                        if (attachment["@odata.type"] === "#microsoft.graph.fileAttachment") {
                             if (!attachment.contentBytes) {
                                context.log.warn(`Attachment "${attachment.name}" (ID: ${attachment.id}) is fileAttachment but contentBytes are missing. Skipping.`);
                                continue;
                            }
                            attachmentData.contentBytes = attachment.contentBytes;
                        } else if (attachment["@odata.type"] === "#microsoft.graph.itemAttachment") {
                             if (!attachment.item) {
                                context.log.warn(`Attachment "${attachment.name}" (ID: ${attachment.id}) is itemAttachment but item data is missing. Skipping.`);
                                continue;
                            }
                            attachmentData.item = attachment.item;
                        } else if (attachment["@odata.type"] === "#microsoft.graph.referenceAttachment") {
                            if(!attachment.sourceUrl || !attachment.providerType) {
                                 context.log.warn(`Attachment "${attachment.name}" (ID: ${attachment.id}) is referenceAttachment but sourceUrl or providerType is missing. Skipping.`);
                                continue;
                            }
                            attachmentData.sourceUrl = attachment.sourceUrl;
                            attachmentData.providerType = attachment.providerType;
                            if (attachment.permission) attachmentData.permission = attachment.permission;
                            if (typeof attachment.isFolder === 'boolean') attachmentData.isFolder = attachment.isFolder;
                        }
                        
                        await client.api(`/me/messages/${draftMessage.id}/attachments`).post(attachmentData);
                        context.log(`Successfully added attachment "${attachment.name}" to draft ${draftMessage.id}.`);
                    } catch (attachError) {
                        const errBody = attachError.body ? JSON.stringify(attachError.body) : 'N/A';
                        context.log.error(`Error adding attachment "${attachment.name}" (ID: ${attachment.id}) to draft ${draftMessage.id}: ${attachError.message}. Error Body: ${errBody}`);
                    }
                }
            }

            context.log(`Sending the new message (draft ID: ${draftMessage.id})...`);
            await client.api(`/me/messages/${draftMessage.id}/send`).post({});
            context.log(`Successfully sent forwarded message. Original message ID was ${messageId}.`);

            context.log(`Moving original message (ID: ${messageId}) to deleted items...`);
            await client.api(`/me/messages/${messageId}/move`).post({ destinationId: "deleteditems" });
            context.log(`Successfully moved original message (ID: ${messageId}) to deleted items.`);

            context.log("Email forwarding process completed successfully.");
            context.res = {
                status: 200,
                body: { success: true, message: "Email forwarded and original moved to deleted items successfully." }
            };
        } catch (error) {
            context.log.error(`Error during message processing/forwarding for message ID "${messageId}": ${error.message}`);
            let errMsg = `Error processing message ID "${messageId}": ${error.message}`;
            if (error.statusCode && error.code) { 
                errMsg = `Graph API Error (${error.code}) for message ID "${messageId}": ${error.message}`;
            }
            context.res = {
                status: (error.statusCode === 404 ? 404 : 500), 
                body: { success: false, error: errMsg, messageIdUsed: messageId }
            };
        }
    } catch (error) {
        context.log.error(`Unhandled error in email forwarding Azure Function: ${error.message}`);
        context.res = {
            status: 500,
            body: { success: false, error: `Critical error in email forwarding process: ${error.message}` }
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
    if (context && context.log) {
        context.log.info(`Attempting to translate EWS ID "${exchangeId}" to REST ID.`);
    }
    try {
        const response = await client.api('/me/translateExchangeIds').post({
            inputIds: [exchangeId],
            targetIdType: "restId",
            sourceIdType: "ewsId"
        });
        if (response && response.value && response.value.length > 0 && response.value[0].targetId) {
            if (context && context.log) context.log.info(`Successfully translated EWS ID "${exchangeId}" to REST ID "${response.value[0].targetId}".`);
            return response.value[0].targetId;
        } else {
            if (context && context.log) context.log.warn(`EWS ID translation for "${exchangeId}" did not return a targetId. Response: ${JSON.stringify(response)}`);
            throw new Error("No translated targetId returned from translateExchangeIds API.");
        }
    } catch (error) {
        const errMsg = error.body ? JSON.stringify(error.body) : error.message;
        if (context && context.log) context.log.error(`EWS ID translation failed for "${exchangeId}": ${errMsg}`);
        throw new Error(`EWS ID to REST ID translation failed: ${errMsg}`);
    }
}
