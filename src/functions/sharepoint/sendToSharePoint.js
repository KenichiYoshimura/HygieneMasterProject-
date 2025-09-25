const axios = require('axios');
const { logMessage, handleError } = require('../utils');

// SharePoint configuration from environment variables
const SHAREPOINT_SITE_URL = process.env.SHAREPOINT_SITE_URL;
const SHAREPOINT_CLIENT_ID = process.env.SHAREPOINT_CLIENT_ID;
const SHAREPOINT_CLIENT_SECRET = process.env.SHAREPOINT_CLIENT_SECRET;
const SHAREPOINT_TENANT_ID = process.env.SHAREPOINT_TENANT_ID;
const SHAREPOINT_DOCUMENT_LIBRARY = process.env.SHAREPOINT_DOCUMENT_LIBRARY || 'Documents';

// Extract site info from URL for Graph API
const siteUrl = new URL(SHAREPOINT_SITE_URL);
const hostname = siteUrl.hostname; // yysolutions.sharepoint.com
const sitePath = siteUrl.pathname; // /sites/ATEMS

// Get access token for Microsoft Graph API
async function getSharePointAccessToken(context) {
    try {
        logMessage(`🔐 Getting Microsoft Graph access token for tenant: ${SHAREPOINT_TENANT_ID}`, context);
        
        const tokenUrl = `https://login.microsoftonline.com/${SHAREPOINT_TENANT_ID}/oauth2/v2.0/token`;
        logMessage(`🔗 Token URL: ${tokenUrl}`, context);
        
        const params = new URLSearchParams();
        params.append('client_id', SHAREPOINT_CLIENT_ID);
        params.append('client_secret', SHAREPOINT_CLIENT_SECRET);
        params.append('scope', 'https://graph.microsoft.com/.default'); // Changed to Graph API scope
        params.append('grant_type', 'client_credentials');
        
        logMessage(`📋 Requesting token with Microsoft Graph scope`, context);

        const response = await axios.post(tokenUrl, params, {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        });

        logMessage(`✅ Microsoft Graph access token obtained successfully`, context);
        return response.data.access_token;
    } catch (error) {
        logMessage(`❌ Microsoft Graph authentication failed: ${error.message}`, context);
        if (error.response) {
            logMessage(`❌ Response status: ${error.response.status}`, context);
            logMessage(`❌ Response data: ${JSON.stringify(error.response.data)}`, context);
        }
        handleError(error, 'Microsoft Graph Authentication', context);
        throw error;
    }
}

// Upload JSON report to SharePoint using Microsoft Graph API
// Simplified upload without retry (since folder is created beforehand)
async function uploadJsonToSharePoint(jsonData, fileName, folderPath, context) {
    try {
        logMessage(`📤 Starting JSON upload: ${fileName}`, context);
        
        const accessToken = await getSharePointAccessToken(context);
        const jsonContent = JSON.stringify(jsonData, null, 2);
        const buffer = Buffer.from(jsonContent, 'utf8');
        logMessage(`📊 JSON buffer size: ${buffer.length} bytes`, context);
        
        const graphUploadUrl = `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}:/drives/root:/${folderPath}/${fileName}:/content`;
        logMessage(`🔗 Upload URL: ${graphUploadUrl}`, context);
        
        const response = await axios.put(graphUploadUrl, buffer, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json',
            },
            timeout: 30000
        });

        logMessage(`✅ JSON uploaded successfully: ${fileName}`, context);
        return response.data;
        
    } catch (error) {
        logMessage(`❌ JSON upload failed: ${fileName} - ${error.message}`, context);
        if (error.response) {
            logMessage(`❌ Response status: ${error.response.status}`, context);
            logMessage(`❌ Response data: ${JSON.stringify(error.response.data)}`, context);
        }
        throw error;
    }
}

// Upload text report to SharePoint using Microsoft Graph API
async function uploadTextToSharePoint(textContent, fileName, folderPath, context) {
    try {
        logMessage(`📤 Starting text upload: ${fileName}`, context);
        
        const accessToken = await getSharePointAccessToken(context);
        const buffer = Buffer.from(textContent, 'utf8');
        logMessage(`📊 Text buffer size: ${buffer.length} bytes`, context);
        
        const graphUploadUrl = `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}:/drives/root:/${folderPath}/${fileName}:/content`;
        
        const response = await axios.put(graphUploadUrl, buffer, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'text/plain; charset=utf-8',
            },
            timeout: 30000
        });

        logMessage(`✅ Text uploaded successfully: ${fileName}`, context);
        return response.data;
    } catch (error) {
        logMessage(`❌ Text upload failed: ${fileName} - ${error.message}`, context);
        throw error;
    }
}

// Upload original document to SharePoint using Microsoft Graph API
async function uploadOriginalDocumentToSharePoint(base64Content, fileName, folderPath, context) {
    try {
        logMessage(`📤 Starting original document upload: ${fileName}`, context);
        
        const accessToken = await getSharePointAccessToken(context);
        const buffer = Buffer.from(base64Content, 'base64');
        logMessage(`📊 Document buffer size: ${buffer.length} bytes`, context);
        
        const graphUploadUrl = `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}:/drives/root:/${folderPath}/${fileName}:/content`;
        
        const response = await axios.put(graphUploadUrl, buffer, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/octet-stream',
            },
            timeout: 60000
        });

        logMessage(`✅ Original document uploaded successfully: ${fileName}`, context);
        return response.data;
    } catch (error) {
        logMessage(`❌ Original document upload failed: ${fileName} - ${error.message}`, context);
        throw error;
    }
}

async function ensureSharePointFolder(folderPath, context) {
    try {
        logMessage(`📁 Ensuring SharePoint folder exists via Graph: ${folderPath}`, context);

        const accessToken = await getSharePointAccessToken(context);
        const siteUrl = new URL(SHAREPOINT_SITE_URL);
        const hostname = siteUrl.hostname;
        const sitePath = siteUrl.pathname;

        // Step 1: Get siteId
        const siteResponse = await axios.get(`https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`, {
            headers: {
                Authorization: `Bearer ${accessToken}`
            }
        });
        const siteId = siteResponse.data.id;
        logMessage(`📋 Site ID: ${siteId}`, context);

        // Step 2: Create folder structure recursively
        const folderParts = folderPath.split('/').filter(part => part);
        let currentPath = '';

        for (const folderName of folderParts) {
            currentPath = currentPath ? `${currentPath}/${folderName}` : folderName;

            // Check if folder exists
            const checkUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${currentPath}`;
            try {
                await axios.get(checkUrl, {
                    headers: {
                        Authorization: `Bearer ${accessToken}`
                    }
                });
                logMessage(`📁 Folder already exists: ${currentPath}`, context);
            } catch (checkError) {
                if (checkError.response && checkError.response.status === 404) {
                    // Create folder
                    const createUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${currentPath.substring(0, currentPath.lastIndexOf('/')) || ''}/children`;
                    logMessage(`📁 Creating folder '${folderName}' at '${createUrl}'`, context);

                    await axios.post(createUrl, {
                        name: folderName,
                        folder: {},
                        "@microsoft.graph.conflictBehavior": "rename"
                    }, {
                        headers: {
                            Authorization: `Bearer ${accessToken}`,
                            'Content-Type': 'application/json'
                        }
                    });

                    logMessage(`✅ Created folder: ${currentPath}`, context);
                } else {
                    logMessage(`❌ Error checking folder: ${currentPath}`, context);
                    throw checkError;
                }
            }
        }

        logMessage(`✅ SharePoint folder structure ensured: ${folderPath}`, context);

    } catch (error) {
        logMessage(`❌ Folder creation via Graph failed: ${error.message}`, context);
        if (error.response) {
            logMessage(`❌ Response status: ${error.response.status}`, context);
            logMessage(`❌ Response data: ${JSON.stringify(error.response.data)}`, context);
        }
        throw error;
    }
}

module.exports = {
    getSharePointAccessToken,
    uploadJsonToSharePoint,
    uploadTextToSharePoint,
    uploadOriginalDocumentToSharePoint,
    ensureSharePointFolder
};
