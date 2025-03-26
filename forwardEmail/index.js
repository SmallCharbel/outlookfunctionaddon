// Required for Microsoft Graph client in Node.js environment
require('isomorphic-fetch');

const { Client } = require('@microsoft/microsoft-graph-client');

module.exports = async function (context, req) {
    context.log('Processing email forwarding request');
    
    try {
        // Log the full request body for debugging
        context.log('Request body:', JSON.stringify(req.body || {}));
        
        // Check if required parameters are present with better validation
        if (!req.body) {
            context.log.error('Request body is missing');
            context.res = {
                status: 400,
                body: { error: 'Request body is missing', success: false }
            };
            return;
        }
        
        const messageId = req.body.messageId;
        context.log(`Received messageId: ${messageId}`);
        
        if (!messageId) {
            context.log.error('Missing messageId parameter');
            context.res = {
                status: 400,
                body: { error: 'Missing messageId parameter', success: false }
            };
            return;
        }

        // Try to get token from authorization header first
        let accessToken = null;
        if (req.headers && req.headers.authorization) {
            context.log('Found authorization header');
            accessToken = req.headers.authorization.replace('Bearer ', '');
        }
        // Fall back to token in request body
        else if (req.body && req.body.accessToken) {
            context.log('Using token from request body');
            accessToken = req.body.accessToken;
        }
        else {
            context.log.error('No authentication token provided');
            context.res = {
                status: 401,
                body: { error: 'Authentication required', success: false }
            };
            return;
        }
        
        context.log('Creating Graph client with delegated token');
        
        // Create Graph client with user's access token
        const client = Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });
        
        // Extract the mailbox address from the request
        const mailboxAddress = req.body.mailboxAddress;
        
        // Try multiple approaches to get the message
        let message = null;
        let messageEndpoint = null;
        
        // Approach 1: Try using the ID directly with proper encoding
        try {
            context.log('Approach 1: Using direct message ID with encoding');
            const encodedMessageId = encodeURIComponent(messageId);
            context.log(`Encoded messageId: ${encodedMessageId}`);
            
            messageEndpoint = mailboxAddress ? 
                `/users/${mailboxAddress}/messages/${encodedMessageId}` : 
                `/me/messages/${encodedMessageId}`;
                
            context.log(`Using API path: ${messageEndpoint}`);
            
            message = await client.api(messageEndpoint)
                .select('subject,body,toRecipients,ccRecipients')
                .get();
                
            context.log('Successfully retrieved message using direct ID');
        } catch (error1) {
            context.log.error(`Approach 1 failed: ${error1.message}`);
            
            // Approach 2: Try using a filter query
            try {
                context.log('Approach 2: Using filter query');
                
                // Extract a unique part of the message ID to use in the filter
                // This assumes the message ID has a format where we can extract a unique identifier
                const idParts = messageId.split('AAA');
                const uniqueIdPart = idParts.length > 1 ? idParts[1].substring(0, 20) : messageId.substring(0, 20);
                
                const filterEndpoint = mailboxAddress ? 
                    `/users/${mailboxAddress}/messages` : 
                    `/me/messages`;
                    
                context.log(`Using filter endpoint: ${filterEndpoint}`);
                context.log(`Filtering with ID part: ${uniqueIdPart}`);
                
                const messages = await client.api(filterEndpoint)
                    .filter(`contains(id,'${uniqueIdPart}')`)
                    .select('id,subject,body,toRecipients,ccRecipients')
                    .top(1)
                    .get();
                    
                if (messages && messages.value && messages.value.length > 0) {
                    message = messages.value[0];
                    messageEndpoint = mailboxAddress ? 
                        `/users/${mailboxAddress}/messages/${message.id}` : 
                        `/me/messages/${message.id}`;
                        
                    context.log('Successfully retrieved message using filter query');
                } else {
                    throw new Error('No messages found matching the filter criteria');
                }
            } catch (error2) {
                context.log.error(`Approach 2 failed: ${error2.message}`);
                
                // Approach 3: Try using the beta endpoint
                try {
                    context.log('Approach 3: Using beta endpoint');
                    
                    // Create a new client specifically for the beta endpoint
                    const betaClient = Client.init({
                        authProvider: (done) => {
                            done(null, accessToken);
                        },
                        baseUrl: "https://graph.microsoft.com/beta"
                    });
                    
                    const encodedMessageId = encodeURIComponent(messageId);
                    const betaEndpoint = mailboxAddress ? 
                        `/users/${mailboxAddress}/messages/${encodedMessageId}` : 
                        `/me/messages/${encodedMessageId}`;
                        
                    context.log(`Using beta endpoint: ${betaEndpoint}`);
                    
                    message = await betaClient.api(betaEndpoint)
                        .select('subject,body,toRecipients,ccRecipients')
                        .get();
                        
                    messageEndpoint = betaEndpoint;
                    context.log('Successfully retrieved message using beta endpoint');
                } catch (error3) {
                    context.log.error(`Approach 3 failed: ${error3.message}`);
                    throw new Error(`Failed to retrieve message using all approaches: ${error1.message}`);
                }
            }
        }
        
        if (!message) {
            throw new Error('Failed to retrieve the message');
        }
        
        // Now that we have the message, continue with the rest of the process
        context.log('Successfully retrieved original message');
        
        // 2. Get message attachments
        context.log('Fetching attachments...');
        const attachmentsResponse = await client.api(`${messageEndpoint}/attachments`).get();
        const attachments = attachmentsResponse.value || [];
        
        // 3. Create a draft of the new message
        context.log('Creating new message draft...');
        const newMessage = {
            subject: message.subject,
            body: {
                contentType: 'html',
                content: message.body.content
            },
            toRecipients: message.toRecipients,
            ccRecipients: message.ccRecipients
        };
        
        const draftApiPath = mailboxAddress ? `/users/${mailboxAddress}/messages` : '/me/messages';
        const draft = await client.api(draftApiPath).post(newMessage);
        
        // 4. Add each attachment to the new message
        if (attachments.length > 0) {
            context.log(`Adding ${attachments.length} attachments...`);
            for (const attachment of attachments) {
                context.log(`Adding attachment: ${attachment.name}`);
                const attachmentApiPath = mailboxAddress ? 
                    `/users/${mailboxAddress}/messages/${draft.id}/attachments` : 
                    `/me/messages/${draft.id}/attachments`;
                    
                await client.api(attachmentApiPath).post({
                    '@odata.type': '#microsoft.graph.fileAttachment',
                    name: attachment.name,
                    contentBytes: attachment.contentBytes,
                    contentType: attachment.contentType
                });
            }
        }
        
        // 5. Send the new message
        context.log('Sending the new message...');
        const sendApiPath = mailboxAddress ? 
            `/users/${mailboxAddress}/messages/${draft.id}/send` : 
            `/me/messages/${draft.id}/send`;
            
        await client.api(sendApiPath).post({});
        
        // 6. Move original to deleted items
        context.log('Moving original message to deleted items...');
        const moveApiPath = mailboxAddress ? 
            `/users/${mailboxAddress}/messages/${message.id}/move` : 
            `/me/messages/${message.id}/move`;
            
        await client.api(moveApiPath).post({
            destinationId: 'deleteditems'
        });
        
        context.log('Process completed successfully');
        context.res = {
            status: 200,
            body: { success: true }
        };
    } catch (error) {
        // Enhanced error logging
        context.log.error(`Error: ${error.message}`);
        
        // Log specific details about Graph API errors
        if (error.statusCode) {
            context.log.error(`Status code: ${error.statusCode}`);
            context.log.error(`Error code: ${error.code}`);
            context.log.error(`Request ID: ${error.requestId}`);
            if (error.body) {
                context.log.error(`Error body: ${error.body}`);
            }
        }
        
        context.log.error(error);
        
        context.res = {
            status: 500,
            body: { 
                error: error.message,
                code: error.code || 'UNKNOWN_ERROR',
                details: error.body || null,
                success: false
            }
        };
    }
};