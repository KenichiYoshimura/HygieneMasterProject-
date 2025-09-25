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

        // Step 1: Get siteId
        const siteResponse = await axios.get(`https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`, {
            headers: {
                Authorization: `Bearer ${accessToken}`
            }
        });
        const siteId = siteResponse.data.id;
        logMessage(`📋 Site ID: ${siteId}`, context);

        // Step 2: Get drive ID
        const driveResponse = await axios.get(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive`, {
            headers: {
                Authorization: `Bearer ${accessToken}`
            }
        });
        const driveId = driveResponse.data.id;
        logMessage(`📋 Drive ID: ${driveId}`, context);

        // Step 3: Create folder structure using drive items API
        const folderParts = folderPath.split('/').filter(part => part);
        let currentItemId = 'root';

        for (const folderName of folderParts) {
            logMessage(`📁 Processing folder: ${folderName}`, context);

            try {
                // First, try to find if folder already exists
                const childrenUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${currentItemId}/children`;
                const childrenResponse = await axios.get(childrenUrl, {
                    headers: {
                        Authorization: `Bearer ${accessToken}`
                    }
                });

                const existingFolder = childrenResponse.data.value.find(
                    item => item.name === folderName && item.folder
                );

                if (existingFolder) {
                    currentItemId = existingFolder.id;
                    logMessage(`📁 Found existing folder '${folderName}' with ID: ${currentItemId}`, context);
                } else {
                    // Create new folder
                    logMessage(`📁 Creating new folder '${folderName}' in item ID: ${currentItemId}`, context);
                    
                    const createUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${currentItemId}/children`;
                    logMessage(`📁 Create URL: ${createUrl}`, context);

                    const createResponse = await axios.post(createUrl, {
                        name: folderName,
                        folder: {},
                        "@microsoft.graph.conflictBehavior": "rename"
                    }, {
                        headers: {
                            Authorization: `Bearer ${accessToken}`,
                            'Content-Type': 'application/json'
                        }
                    });

                    currentItemId = createResponse.data.id;
                    logMessage(`✅ Created folder '${folderName}' with ID: ${currentItemId}`, context);
                }

            } catch (folderError) {
                logMessage(`❌ Error processing folder '${folderName}': ${folderError.message}`, context);
                if (folderError.response) {
                    logMessage(`❌ Folder error status: ${folderError.response.status}`, context);
                    logMessage(`❌ Folder error data: ${JSON.stringify(folderError.response.data)}`, context);
                }
                throw folderError;
            }
        }

        logMessage(`✅ All folders ensured successfully. Final path: ${folderPath}`, context);
        return true;

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
