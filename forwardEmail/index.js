// Required for Microsoft Graph client in Node.js environment
require('isomorphic-fetch');

const { Client } = require('@microsoft/microsoft-graph-client');

module.exports = async function (context, req) {
    context.log('Processing email forwarding request');

    try {
        // Check if required parameters are present
        const messageId = req.body && req.body.messageId;
        
        if (!messageId) {
            context.log.error('Missing messageId parameter');
            context.res = {
                status: 400,
                body: { error: 'Missing messageId parameter.', success: false }
            };
            return;
        }

        // Check for auth header
        if (!req.headers || !req.headers.authorization) {
            context.log.error('No authorization header provided');
            context.res = {
                status: 401,
                body: { error: 'Authentication required', success: false }
            };
            return;
        }

        // Extract the token from the Authorization header
        const authHeader = req.headers.authorization;
        const accessToken = authHeader.replace('Bearer ', '');
        
        context.log('Creating Graph client with delegated token');
        
        // Create Graph client with user's access token
        const client = Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });
        
        // 1. Get the original message
        context.log('Fetching original message...');
        const message = await client.api(`/me/messages/${messageId}`)
            .select('subject,body,toRecipients,ccRecipients')
            .get();
        
        // 2. Get message attachments
        context.log('Fetching attachments...');
        const attachmentsResponse = await client.api(`/me/messages/${messageId}/attachments`).get();
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
        
        const draft = await client.api('/me/messages').post(newMessage);
        
        // 4. Add each attachment to the new message
        if (attachments.length > 0) {
            context.log(`Adding ${attachments.length} attachments...`);
            for (const attachment of attachments) {
                context.log(`Adding attachment: ${attachment.name}`);
                await client.api(`/me/messages/${draft.id}/attachments`).post({
                    '@odata.type': '#microsoft.graph.fileAttachment',
                    name: attachment.name,
                    contentBytes: attachment.contentBytes,
                    contentType: attachment.contentType
                });
            }
        }
        
        // 5. Send the new message
        context.log('Sending the new message...');
        await client.api(`/me/messages/${draft.id}/send`).post({});
        
        // 6. Move original to deleted items
        context.log('Moving original message to deleted items...');
        await client.api(`/me/messages/${messageId}/move`).post({
            destinationId: 'deleteditems'
        });
        
        context.log('Process completed successfully');
        context.res = {
            status: 200,
            body: { success: true }
        };
    } catch (error) {
        context.log.error(`Error: ${error.message}`);
        context.log.error(error);
        
        context.res = {
            status: 500,
            body: { 
                error: error.message,
                success: false
            }
        };
    }
};