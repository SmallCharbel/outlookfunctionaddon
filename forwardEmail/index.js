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
        // Parse request body
        const body = req.body || {};
        context.log(`Request body: ${JSON.stringify(body)}`);
        
        const { messageId, accessToken } = body;
        
        if (!messageId || !accessToken) {
            context.res = {
                status: 400,
                body: "Missing required parameters: messageId and accessToken"
            };
            return;
        }
        
        context.log(`Received messageId: ${messageId}`);
        
        // Check for authorization header
        if (!accessToken) {
            context.log.error("No authorization token provided");
            context.res = {
                status: 401,
                body: "Unauthorized: No token provided"
            };
            return;
        }
        
        context.log("Found authorization header");
        
        // Create Microsoft Graph client
        const client = getAuthenticatedClient(accessToken);
        context.log("Creating Graph client with delegated token");
        
        // Get the message - using a different approach
        try {
            // Try a different approach that doesn't use the message ID directly in the URL path
            context.log("Trying to get message with filter approach");
            
            // First, try to get messages with a filter
            const messages = await client.api('/me/messages')
                .top(10)  // Limit to 10 messages for performance
                .get();
            
            context.log(`Retrieved ${messages.value.length} messages`);
            
            // Find the message with the matching ID
            const message = messages.value.find(msg => msg.id === messageId);
            
            if (!message) {
                throw new Error("Message not found in recent messages");
            }
            
            context.log("Message found successfully");
            context.log(`Subject: ${message.subject}`);
            
            // Return success for now - you can add forwarding logic later
            context.res = {
                status: 200,
                body: { 
                    success: true, 
                    message: "Message found successfully",
                    subject: message.subject
                }
            };
        } catch (error) {
            context.log.error(`Error retrieving message: ${error.message}`);
            context.res = {
                status: 500,
                body: `Error retrieving message: ${error.message}`
            };
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