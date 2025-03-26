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
        
        // 1. Get the original message - FIXED API FORMAT
        // The error is happening because messageId might be malformed or not properly encoded
        context.log(`Fetching original message with ID: ${messageId}`);
        
        // Try to fetch the message using properly formatted endpoint
        const mailboxAddress = req.body.mailboxAddress;
        let apiPath;
        
        if (mailboxAddress) {
            // If mailbox address is provided, use it in the path
            context.log(`Using mailbox address: ${mailboxAddress}`);
            apiPath = `/users/${mailboxAddress}/messages/${messageId}`;
        } else {
            // Otherwise use /me endpoint
            apiPath = `/me/messages/${messageId}`;
        }
        
        context.log(`Using API path: ${apiPath}`);
        
        const message = await client.api(apiPath)
            .select('subject,body,toRecipients,ccRecipients')
            .get();
        
        context.log('Successfully retrieved original message');
        
        // 2. Get message attachments
        context.log('Fetching attachments...');
        const attachmentsResponse = await client.api(`${apiPath}/attachments`).get();
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
            `/users/${mailboxAddress}/messages/${messageId}/move` : 
            `/me/messages/${messageId}/move`;
            
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