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
async function uploadJsonToSharePoint(jsonData, fileName, folderPath, context) {
    try {
        logMessage(`📤 Starting JSON upload via Microsoft Graph: ${fileName}`, context);
        logMessage(`📁 Target folder: ${folderPath}`, context);
        
        const accessToken = await getSharePointAccessToken(context);
        
        const jsonContent = JSON.stringify(jsonData, null, 2);
        const buffer = Buffer.from(jsonContent, 'utf8');
        logMessage(`📊 JSON buffer size: ${buffer.length} bytes`, context);
        
        // Microsoft Graph API endpoint for file upload
        const graphUploadUrl = `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}:/drives/root:/${folderPath}/${fileName}:/content`;
        logMessage(`🔗 Graph Upload URL: ${graphUploadUrl}`, context);
        
        logMessage(`📤 Sending request to Microsoft Graph...`, context);
        const response = await axios.put(graphUploadUrl, buffer, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json',
            },
            timeout: 30000
        });

        logMessage(`✅ JSON uploaded via Microsoft Graph successfully: ${fileName}`, context);
        logMessage(`📊 Response status: ${response.status}`, context);
        return response.data;
    } catch (error) {
        logMessage(`❌ JSON upload via Graph failed for: ${fileName}`, context);
        logMessage(`❌ Error message: ${error.message}`, context);
        if (error.response) {
            logMessage(`❌ Response status: ${error.response.status}`, context);
            logMessage(`❌ Response data: ${JSON.stringify(error.response.data)}`, context);
        }
        handleError(error, 'Microsoft Graph JSON Upload', context);
        throw error;
    }
}

// Upload text report to SharePoint using Microsoft Graph API
async function uploadTextToSharePoint(textContent, fileName, folderPath, context) {
    try {
        logMessage(`📤 Starting text upload via Microsoft Graph: ${fileName}`, context);
        
        const accessToken = await getSharePointAccessToken(context);
        const buffer = Buffer.from(textContent, 'utf8');
        logMessage(`📄 Text buffer size: ${buffer.length} bytes`, context);
        
        const graphUploadUrl = `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}:/drives/root:/${folderPath}/${fileName}:/content`;
        logMessage(`🔗 Graph Text Upload URL: ${graphUploadUrl}`, context);
        
        const response = await axios.put(graphUploadUrl, buffer, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'text/plain; charset=utf-8',
            },
            timeout: 30000
        });

        logMessage(`✅ Text uploaded via Microsoft Graph successfully: ${fileName}`, context);
        return response.data;
    } catch (error) {
        logMessage(`❌ Text upload via Graph failed for: ${fileName}`, context);
        logMessage(`❌ Error message: ${error.message}`, context);
        if (error.response) {
            logMessage(`❌ Response status: ${error.response.status}`, context);
            logMessage(`❌ Response data: ${JSON.stringify(error.response.data)}`, context);
        }
        handleError(error, 'Microsoft Graph Text Upload', context);
        throw error;
    }
}

// Upload original document to SharePoint using Microsoft Graph API
async function uploadOriginalDocumentToSharePoint(base64Content, fileName, folderPath, context) {
    try {
        logMessage(`📤 Starting original document upload via Microsoft Graph: ${fileName}`, context);
        
        const accessToken = await getSharePointAccessToken(context);
        const buffer = Buffer.from(base64Content, 'base64');
        logMessage(`📊 Original document buffer size: ${buffer.length} bytes`, context);
        
        const graphUploadUrl = `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}:/drives/root:/${folderPath}/${fileName}:/content`;
        logMessage(`🔗 Graph Original Document Upload URL: ${graphUploadUrl}`, context);
        
        const response = await axios.put(graphUploadUrl, buffer, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/octet-stream',
            },
            timeout: 60000
        });

        logMessage(`✅ Original document uploaded via Microsoft Graph successfully: ${fileName}`, context);
        return response.data;
    } catch (error) {
        logMessage(`❌ Original document upload via Graph failed for: ${fileName}`, context);
        logMessage(`❌ Error message: ${error.message}`, context);
        if (error.response) {
            logMessage(`❌ Response status: ${error.response.status}`, context);
            logMessage(`❌ Response data: ${JSON.stringify(error.response.data)}`, context);
        }
        handleError(error, 'Microsoft Graph Original Document Upload', context);
        throw error;
    }
}

// Create SharePoint folder using Microsoft Graph API
async function ensureSharePointFolder(folderPath, context) {
    try {
        logMessage(`📁 Ensuring SharePoint folder exists via Graph: ${folderPath}`, context);
        
        const accessToken = await getSharePointAccessToken(context);
        
        // Create nested folders one by one
        const folderParts = folderPath.split('/').filter(part => part);
        let currentPath = '';
        
        for (const folderName of folderParts) {
            currentPath += `/${folderName}`;
            
            try {
                const folderCreateUrl = `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}:/drives/root${currentPath.substring(0, currentPath.lastIndexOf('/'))}:/children`;
                
                await axios.post(folderCreateUrl, {
                    name: folderName,
                    folder: {}
                }, {
                    headers: {
                        'Authorization': `Bearer ${accessToken}`,
                        'Content-Type': 'application/json'
                    }
                });
                
                logMessage(`📁 Created folder: ${currentPath}`, context);
            } catch (folderError) {
                if (folderError.response && folderError.response.status === 409) {
                    logMessage(`📁 Folder already exists: ${currentPath}`, context);
                } else {
                    logMessage(`⚠️ Could not create folder ${currentPath}: ${folderError.message}`, context);
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
        // Don't throw error for folder creation failures
        logMessage(`⚠️ Continuing without folder creation...`, context);
    }
}

module.exports = {
    getSharePointAccessToken,
    uploadJsonToSharePoint,
    uploadTextToSharePoint,
    uploadOriginalDocumentToSharePoint,
    ensureSharePointFolder
};
