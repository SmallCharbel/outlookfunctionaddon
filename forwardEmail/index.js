// Required for Microsoft Graph client in Node.js environment
require('isomorphic-fetch');

const { Client } = require('@microsoft/microsoft-graph-client');

// --- searchMessageByMetadata: Fallback search logic ---
// This function is used if ewsItemId is not available or conversion fails.
// It relies heavily on receivedTime for the initial Graph query.
async function searchMessageByMetadata(client, subjectFromRequest, recipientsFromRequest, receivedTimeFromRequest, context) {
    const encodeOData = (str) => str ? str.replace(/'/g, "''") : "";

    if (!receivedTimeFromRequest || typeof receivedTimeFromRequest !== 'string' || receivedTimeFromRequest.trim() === '') {
        if (context && context.log) {
            context.log.error("searchMessageByMetadata: receivedTimeFromRequest is missing, not a string, or empty. It is mandatory for a precise fallback search.");
        }
        throw new Error("receivedTimeFromRequest must be a valid ISO date string and is required for metadata fallback search.");
    }

    const initialODataFilter = `receivedDateTime eq ${encodeOData(receivedTimeFromRequest)}`;
    if (context && context.log) context.log.info(`searchMessageByMetadata: Constructing OData filter for Graph API: ${initialODataFilter}`);

    try {
        const response = await client.api('/me/messages')
            .filter(initialODataFilter)
            .select('id,receivedDateTime,subject,toRecipients')
            .top(5)
            .get();

        if (response.value && response.value.length > 0) {
            if (context && context.log) context.log.info(`searchMessageByMetadata: Initial query by receivedTime ("${receivedTimeFromRequest}") returned ${response.value.length} message(s). Performing detailed validation...`);
            
            const fullRecipientList = recipientsFromRequest ? recipientsFromRequest.split(';').map(r => r.trim().toLowerCase()).filter(r => r) : [];

            for (const message of response.value) {
                let subjectMatch = !subjectFromRequest;
                if (subjectFromRequest) {
                    const messageSubjectNormalized = message.subject ? message.subject.trim().toLowerCase() : "";
                    const requestSubjectNormalized = subjectFromRequest.trim().toLowerCase();
                    subjectMatch = (messageSubjectNormalized === requestSubjectNormalized);
                    if (!subjectMatch && context && context.log) context.log.info(`searchMessageByMetadata: Message ${message.id} (Received: ${message.receivedDateTime}): Subject mismatch. Expected: "${requestSubjectNormalized}", Actual: "${messageSubjectNormalized}".`);
                }
                if (!subjectMatch) continue;

                let recipientsMatch = fullRecipientList.length === 0;
                if (fullRecipientList.length > 0) {
                    if (message.toRecipients && message.toRecipients.length > 0) {
                        const messageRecipientsSet = new Set(message.toRecipients.map(r => r.emailAddress && r.emailAddress.address ? r.emailAddress.address.toLowerCase() : null).filter(Boolean));
                        let allExpectedRecipientsFound = true;
                        for (const expectedRecipient of fullRecipientList) {
                            if (!messageRecipientsSet.has(expectedRecipient)) {
                                allExpectedRecipientsFound = false;
                                if (context && context.log) context.log.info(`searchMessageByMetadata: Message ${message.id} (Received: ${message.receivedDateTime}): Recipient mismatch. Expected "${expectedRecipient}" not found. Message recipients: [${Array.from(messageRecipientsSet).join(', ')}].`);
                                break;
                            }
                        }
                        if (allExpectedRecipientsFound) recipientsMatch = true;
                    } else if (context && context.log) {
                        context.log.info(`searchMessageByMetadata: Message ${message.id} (Received: ${message.receivedDateTime}): Recipient mismatch. Expected ${fullRecipientList.length} recipients, but message has none.`);
                    }
                }
                if (!recipientsMatch) continue;

                if (context && context.log) context.log.info(`searchMessageByMetadata: SUCCESS - Message ${message.id} (Received: "${message.receivedDateTime}", Subject: "${message.subject}") passed all client-side validations.`);
                return message.id;
            }
            if (context && context.log) context.log.info("searchMessageByMetadata: No messages passed detailed client-side validation after initial query by receivedTime.");
        } else if (context && context.log) {
            context.log.info(`searchMessageByMetadata: Initial OData query by receivedTime ("${receivedTimeFromRequest}") returned no messages.`);
        }
        return null;
    } catch (err) {
        if (context && context.log) context.log.error(`searchMessageByMetadata: Error during Graph API call (Filter: "${initialODataFilter}"): ${err.message}`);
        throw new Error(`Error in searchMessageByMetadata: ${err.message}`);
    }
}

// --- Main function handler ---
module.exports = async function (context, req) {
    context.log("Processing email forwarding request...");

    try {
        context.log(`Request Headers: ${JSON.stringify(req.headers)}`);
        context.log(`Request Body: ${JSON.stringify(req.body || {})}`); // Log the entire body

        const authHeader = req.headers.authorization || '';
        if (!authHeader.startsWith('Bearer ')) {
            context.log.error("Unauthorized: No authorization token provided.");
            context.res = { status: 401, body: { success: false, error: "Unauthorized: No token provided" } };
            return;
        }
        const accessToken = authHeader.substring(7);
        const client = getAuthenticatedClient(accessToken);
        context.log("Graph client created with delegated token.");

        const {
            ewsItemId,              // Preferred: EWS Item ID from Outlook Add-in
            subject,                // For metadata search fallback / verification
            recipients,             // For metadata search fallback / verification (ensure this is "" if no recipients, not undefined/null string)
            receivedTime,           // For metadata search fallback / verification
            useMetadataSearch = true // Default to true to allow fallback, client can set to false
        } = req.body || {};

        let messageIdToProcess = null; // This will hold the Graph REST ID

        // Priority 1: Use ewsItemId if provided
        if (ewsItemId && typeof ewsItemId === 'string' && ewsItemId.trim() !== '') {
            context.log(`EWS Item ID provided: "${ewsItemId}". Attempting conversion to Graph REST ID.`);
            try {
                messageIdToProcess = await convertExchangeId(client, ewsItemId, context);
                if (messageIdToProcess) {
                    context.log(`Successfully converted EWS ID "${ewsItemId}" to Graph REST ID: "${messageIdToProcess}".`);
                } else {
                    // convertExchangeId should ideally throw if it can't return an ID.
                    throw new Error("convertExchangeId returned a non-error falsy value.");
                }
            } catch (conversionError) {
                context.log.error(`Error converting EWS ID "${ewsItemId}": ${conversionError.message}.`);
                if (!useMetadataSearch) {
                    context.log.error("EWS ID conversion failed and metadata search fallback is disabled.");
                    context.res = { status: 500, body: { success: false, error: `Failed to resolve message ID from EWS ID: ${conversionError.message}. Metadata search fallback is disabled.` } };
                    return;
                }
                context.log.warn("Falling back to metadata search due to EWS ID conversion failure.");
                messageIdToProcess = null; // Ensure it's null for the next check
            }
        } else {
            context.log("EWS Item ID not provided or empty. Will attempt metadata search if enabled.");
        }

        // Priority 2: Fallback to metadata search if ewsItemId was not used/failed AND useMetadataSearch is true
        if (!messageIdToProcess && useMetadataSearch) {
            context.log.info(`Attempting metadata search. Criteria: Subject="${subject}", Recipients="${recipients}", ReceivedTime="${receivedTime}"`);
            // Defensive check for recipients string value before logging/using
            const recipientsForLog = typeof recipients === 'string' ? recipients : JSON.stringify(recipients);
            try {
                messageIdToProcess = await searchMessageByMetadata(client, subject, recipients, receivedTime, context);
                if (messageIdToProcess) {
                    context.log(`Metadata search successful. Found message with Graph REST ID: "${messageIdToProcess}".`);
                } else {
                    context.log.error(`Metadata search did not find a matching message. Criteria: Subject="${subject}", Recipients="${recipientsForLog}", ReceivedTime="${receivedTime}"`);
                    context.res = { status: 404, body: { success: false, error: `No message found via metadata search matching criteria (S: "${subject}", R: "${recipientsForLog}", T: "${receivedTime}")` } };
                    return;
                }
            } catch (searchError) {
                context.log.error(`Error during metadata search: ${searchError.message}`);
                context.res = { status: 500, body: { success: false, error: `Error searching for message via metadata: ${searchError.message}` } };
                return;
            }
        }

        // Final check: If no messageId could be determined
        if (!messageIdToProcess) {
            context.log.error("Failed to determine a valid message ID through any method (EWS ID conversion or Metadata Search).");
            context.res = { status: 400, body: { success: false, error: "Unable to identify the target message. Ensure EWS Item ID is provided or metadata search criteria are correct and search is enabled." } };
            return;
        }

        context.log(`Proceeding to process message with Graph REST ID: "${messageIdToProcess}"`);
        // --- Message Processing (Fetch, Draft, Attachments, Send, Move) ---
        try {
            const message = await client.api(`/me/messages/${messageIdToProcess}`)
                .select('subject,body,toRecipients,ccRecipients,bccRecipients,from,hasAttachments,importance,isRead')
                .get();
            context.log(`Successfully retrieved original message: "${message.subject}" (ID: ${message.id})`);

            let attachments = [];
            if (message.hasAttachments) {
                context.log("Original message has attachments. Fetching attachment details...");
                const attachmentsResponse = await client.api(`/me/messages/${messageIdToProcess}/attachments`).get();
                attachments = attachmentsResponse.value || [];
                context.log(`Found ${attachments.length} attachments.`);
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
            context.log(`New draft message created with ID: ${draftMessage.id}.`);

            if (attachments.length > 0) {
                context.log(`Adding ${attachments.length} attachments to the new draft...`);
                for (const attachment of attachments) {
                    try {
                        const attachmentData = {
                            "@odata.type": attachment["@odata.type"],
                            name: attachment.name,
                            contentType: attachment.contentType,
                        };
                        if (attachment["@odata.type"] === "#microsoft.graph.fileAttachment" && attachment.contentBytes) {
                            attachmentData.contentBytes = attachment.contentBytes;
                        } else if (attachment["@odata.type"] === "#microsoft.graph.itemAttachment" && attachment.item) {
                            attachmentData.item = attachment.item;
                        } else if (attachment["@odata.type"] === "#microsoft.graph.referenceAttachment" && attachment.sourceUrl && attachment.providerType) {
                            attachmentData.sourceUrl = attachment.sourceUrl;
                            attachmentData.providerType = attachment.providerType;
                            if (attachment.permission) attachmentData.permission = attachment.permission;
                            if (typeof attachment.isFolder === 'boolean') attachmentData.isFolder = attachment.isFolder;
                        } else if (attachment["@odata.type"] !== "#microsoft.graph.fileAttachment" && attachment["@odata.type"] !== "#microsoft.graph.itemAttachment" && attachment["@odata.type"] !== "#microsoft.graph.referenceAttachment") {
                             context.log.warn(`Unsupported attachment type or missing data for attachment "${attachment.name}" (Type: ${attachment["@odata.type"]}). Skipping.`);
                             continue;
                        } else if(!attachmentData.contentBytes && !attachmentData.item && !attachmentData.sourceUrl) { // If it's a known type but data is missing
                            context.log.warn(`Attachment "${attachment.name}" (Type: ${attachment["@odata.type"]}) is missing required data (e.g. contentBytes, item, sourceUrl). Skipping.`);
                            continue;
                        }
                        await client.api(`/me/messages/${draftMessage.id}/attachments`).post(attachmentData);
                        context.log(`Successfully added attachment "${attachment.name}".`);
                    } catch (attachError) {
                        const errBody = attachError.body ? JSON.stringify(attachError.body) : 'N/A';
                        context.log.error(`Error adding attachment "${attachment.name}": ${attachError.message}. Body: ${errBody}`);
                    }
                }
            }

            context.log(`Sending the new message (draft ID: ${draftMessage.id})...`);
            await client.api(`/me/messages/${draftMessage.id}/send`).post({});
            context.log(`Successfully sent forwarded message. Original message ID was ${messageIdToProcess}.`);

            context.log(`Moving original message (ID: ${messageIdToProcess}) to deleted items...`);
            await client.api(`/me/messages/${messageIdToProcess}/move`).post({ destinationId: "deleteditems" });
            context.log(`Successfully moved original message (ID: ${messageIdToProcess}) to deleted items.`);

            context.log("Email forwarding process completed successfully.");
            context.res = { status: 200, body: { success: true, message: "Email forwarded and original moved to deleted items successfully." } };

        } catch (processingError) {
            context.log.error(`Error during message processing/forwarding for Graph REST ID "${messageIdToProcess}": ${processingError.message}`);
            let errMsg = `Error processing message (ID: "${messageIdToProcess}"): ${processingError.message}`;
            if (processingError.statusCode && processingError.code) {
                errMsg = `Graph API Error (${processingError.code}) for message ID "${messageIdToProcess}": ${processingError.message}`;
            }
            context.res = { status: (processingError.statusCode === 404 ? 404 : 500), body: { success: false, error: errMsg, messageIdUsed: messageIdToProcess } };
        }
    } catch (error) { // Outermost catch
        context.log.error(`Unhandled error in Azure Function: ${error.message}`);
        context.res = { status: 500, body: { success: false, error: `Critical error in email forwarding process: ${error.message}` } };
    }
};

// --- Helper Functions ---
function getAuthenticatedClient(accessToken) {
    const client = Client.init({ authProvider: (done) => done(null, accessToken) });
    return client;
}

async function convertExchangeId(client, ewsItemId, context) {
    if (context && context.log) context.log.info(`convertExchangeId: Attempting to translate EWS ID "${ewsItemId}" to REST ID.`);
    try {
        const response = await client.api('/me/translateExchangeIds').post({
            inputIds: [ewsItemId],
            targetIdType: "restId",
            sourceIdType: "ewsId"
        });
        if (response && response.value && response.value.length > 0 && response.value[0].targetId) {
            if (context && context.log) context.log.info(`convertExchangeId: Successfully translated EWS ID "${ewsItemId}" to REST ID "${response.value[0].targetId}".`);
            return response.value[0].targetId;
        } else {
            if (context && context.log) context.log.warn(`convertExchangeId: Translation for "${ewsItemId}" did not return a targetId. Response: ${JSON.stringify(response)}`);
            throw new Error("No translated targetId returned from translateExchangeIds API for EWS ID: " + ewsItemId);
        }
    } catch (error) {
        const errMsg = error.body ? JSON.stringify(error.body) : error.message;
        if (context && context.log) context.log.error(`convertExchangeId: Translation failed for "${ewsItemId}": ${errMsg}`);
        throw new Error(`EWS ID to REST ID translation failed for "${ewsItemId}": ${errMsg}`);
    }
}

// Note: The old 'validateMessageId' function is no longer needed in the main flow
// as we explicitly handle EWS ID conversion and expect Graph REST IDs otherwise.
