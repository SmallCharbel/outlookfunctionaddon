// Required for Microsoft Graph client in Node.js environment
require('isomorphic-fetch');

const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

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
    context.log('Forward Email function processing request.');
    
    try {
        // Check if we have the necessary data
        if (!req.body) {
            context.log.error('No request body provided');
            context.res = {
                status: 400,
                body: { success: false, error: "No request body provided" }
            };
            return;
        }
        
        // Get the access token from the Authorization header
        let accessToken = null;
        if (req.headers && req.headers.authorization) {
            const authHeader = req.headers.authorization;
            if (authHeader.startsWith('Bearer ')) {
                accessToken = authHeader.substring(7);
                context.log('Access token found in Authorization header');
            }
        }
        
        // Fallback to access token in body if not in header
        if (!accessToken && req.body.accessToken) {
            accessToken = req.body.accessToken;
            context.log('Access token found in request body');
        }
        
        if (!accessToken) {
            context.log.error('No access token provided');
            context.res = {
                status: 401,
                body: { success: false, error: "No access token provided" }
            };
            return;
        }
        
        // Initialize Microsoft Graph client with the provided token
        const authProvider = (callback) => {
            callback(null, accessToken);
        };
        
        const client = Client.init({
            authProvider: authProvider
        });
        
        // Get message ID - either directly provided or search by subject
        let messageId = null;
        
        // Check if we should use subject-based search
        if (req.body.useSubjectSearch && req.body.subject) {
            context.log(`Searching for email with subject: ${req.body.subject}`);
            
            try {
                // Search for the email by subject
                const messages = await client.api('/me/messages')
                    .filter(`subject eq '${req.body.subject.replace(/'/g, "''")}'`)
                    .top(5)
                    .select('id,subject,receivedDateTime')
                    .orderBy('receivedDateTime DESC')
                    .get();
                
                if (messages.value && messages.value.length > 0) {
                    // Use the most recent message with matching subject
                    messageId = messages.value[0].id;
                    context.log(`Found message with ID: ${messageId}`);
                } else {
                    context.log.error(`No messages found with subject: ${req.body.subject}`);
                    context.res = {
                        status: 404,
                        body: {
                            success: false,
                            error: `No messages found with subject: ${req.body.subject}`
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
        } else if (req.body.messageId) {
            // Use the provided message ID
            messageId = req.body.messageId;
            context.log(`Using provided message ID: ${messageId}`);
        } else {
            context.log.error('No message ID or subject provided');
            context.res = {
                status: 400,
                body: {
                    success: false,
                    error: "No message ID or subject provided"
                }
            };
            return;
        }
        
        // Get the forward-to email address from environment variables
        const forwardToEmail = process.env.FORWARD_TO_EMAIL;
        if (!forwardToEmail) {
            context.log.error('Forward-to email address not configured');
            context.res = {
                status: 500,
                body: {
                    success: false,
                    error: "Forward-to email address not configured"
                }
            };
            return;
        }
        
        // Forward the email
        context.log(`Forwarding message ${messageId} to ${forwardToEmail}`);
        
        const forwardRequest = {
            message: {
                toRecipients: [
                    {
                        emailAddress: {
                            address: forwardToEmail
                        }
                    }
                ]
            },
            comment: "Forwarded by Email Forward Add-in"
        };
        
        await client.api(`/me/messages/${messageId}/forward`)
            .post(forwardRequest);
        
        context.log('Email forwarded successfully');
        context.res = {
            status: 200,
            body: {
                success: true,
                message: "Email forwarded successfully"
            }
        };
        
    } catch (error) {
        context.log.error(`Error forwarding email: ${error.message}`);
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