// Required for Microsoft Graph client in Node.js environment
require('isomorphic-fetch');

const { Client } = require('@microsoft/microsoft-graph-client');

// Modify the message retrieval approach
async function getMessageWithRetry(client, messageId) {
    // Try multiple approaches to retrieve the message
    try {
        // Approach 1: Direct access with proper encoding
        console.log("Approach 1: Using direct message ID with proper encoding");
        // Don't re-encode an already encoded messageId
        return await client.api(`/me/messages/${messageId}`).get();
    } catch (error1) {
        console.log(`Approach 1 failed: ${error1.message}`);
        
        try {
            // Approach 2: Try with beta endpoint
            console.log("Approach 2: Using beta endpoint");
            return await client.api(`/beta/me/messages/${messageId}`).get();
        } catch (error2) {
            console.log(`Approach 2 failed: ${error2.message}`);
            
            try {
                // Approach 3: Try with $select to get minimal data
                console.log("Approach 3: Using $select to get minimal data");
                return await client.api(`/me/messages/${messageId}`)
                    .select('id,subject,body,toRecipients,from')
                    .get();
            } catch (error3) {
                console.log(`Approach 3 failed: ${error3.message}`);
                
                // All approaches failed
                throw new Error(`Failed to retrieve message using all approaches: ${error1.message}`);
            }
        }
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
        
        // Get message ID from request body or query parameters
        let messageId = (req.body && req.body.messageId) || 
                       (req.query && req.query.messageId);
        
        context.log(`Original message ID from request: ${messageId || 'Not provided'}`);
        
        if (!messageId) {
            // If no message ID provided, list recent messages to help debugging
            context.log("No message ID provided, retrieving recent messages");
            
            const messages = await client.api('/me/messages')
                .top(5)
                .select('id,subject,receivedDateTime')
                .orderBy('receivedDateTime DESC')
                .get();
            
            context.log(`Retrieved ${messages.value.length} recent messages`);
            messages.value.forEach((msg, index) => {
                context.log(`Message ${index + 1}: ID=${msg.id}, Subject=${msg.subject}`);
            });
            
            context.res = {
                status: 400,
                body: {
                    error: "Missing required parameter: messageId",
                    recentMessages: messages.value.map(m => ({ id: m.id, subject: m.subject }))
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
                .select('subject,body,toRecipients,ccRecipients,from,hasAttachments')
                .get();
            
            context.log(`Successfully retrieved message: ${message.subject}`);
            
            // Fetch attachments if any
            let attachments = [];
            if (message.hasAttachments) {
                context.log("Fetching attachments...");
                const attachmentsResponse = await client.api(`/me/messages/${messageId}/attachments`)
                    .get();
                attachments = attachmentsResponse.value;
                context.log(`Found ${attachments.length} attachments`);
            }
            
            // Create a new message draft
            context.log("Creating new message draft...");
            const newMessage = {
                subject: `FW: ${message.subject}`,
                body: {
                    contentType: message.body.contentType,
                    content: message.body.content
                },
                toRecipients: [
                    {
                        emailAddress: {
                            address: "jmarshall@ptai.com" // Replace with your forwarding address
                        }
                    }
                ]
            };
            
            // Create the draft message
            const draftMessage = await client.api('/me/messages')
                .post(newMessage);
            
            // Add attachments if any
            if (attachments.length > 0) {
                context.log(`Adding ${attachments.length} attachments...`);
                
                for (const attachment of attachments) {
                    context.log(`Adding attachment: ${attachment.name}`);
                    
                    await client.api(`/me/messages/${draftMessage.id}/attachments`)
                        .post({
                            "@odata.type": attachment["@odata.type"],
                            name: attachment.name,
                            contentType: attachment.contentType,
                            contentBytes: attachment.contentBytes
                        });
                }
            }
            
            // Send the message
            context.log("Sending the new message...");
            await client.api(`/me/messages/${draftMessage.id}/send`)
                .post({});
            
            // Move original message to deleted items
            context.log("Moving original message to deleted items...");
            await client.api(`/me/messages/${messageId}/move`)
                .post({
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
            context.log.error(`Error accessing message: ${error.message}`);
            
            // If we get a 404 or other error, try to list recent messages as a fallback
            try {
                context.log("Error retrieving message. Listing recent messages to help troubleshoot...");
                const messages = await client.api('/me/messages')
                    .top(5)
                    .select('id,subject,receivedDateTime')
                    .orderBy('receivedDateTime DESC')
                    .get();
                
                context.res = {
                    status: 404,
                    body: {
                        error: `Message not found: ${error.message}`,
                        messageIdUsed: messageId,
                        recentMessages: messages.value.map(m => ({ id: m.id, subject: m.subject }))
                    }
                };
            } catch (listError) {
                context.res = {
                    status: 500,
                    body: `Error accessing message: ${error.message}`
                };
            }
        }
    } catch (error) {
        context.log.error(`Error: ${error.message}`);
        context.res = {
            status: 500,
            body: `Error forwarding email: ${error.message}`
        };
    }
};

// Helper function to create authenticated client
function getAuthenticatedClient(accessToken) {
    const { Client } = require('@microsoft/microsoft-graph-client');
    
    // Initialize Graph client
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
    
    // Check if ID is too short or just numeric (likely invalid)
    if (id.length < 10 || /^\d+$/.test(id)) {
        return false;
    }
    
    // Check for common issues like stray '/' characters
    if (id.includes('/')) {
        return false;
    }
    
    return true;
}

// Convert Exchange ID to REST format using translateExchangeIds
async function convertExchangeId(client, exchangeId) {
    try {
        const response = await client.api('/me/translateExchangeIds')
            .post({
                inputIds: [exchangeId],
                targetIdType: "restId",
                sourceIdType: "ewsId"
            });
        
        if (response && response.value && response.value.length > 0) {
            return response.value[0].targetId;
        }
        throw new Error("No translated ID returned");
    } catch (error) {
        throw new Error(`Translation failed: ${error.message}`);
    }
}