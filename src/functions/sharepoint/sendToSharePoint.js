const axios = require('axios');
const { logMessage, handleError } = require('../utils');

// SharePoint configuration from environment variables
const SHAREPOINT_SITE_URL = process.env.SHAREPOINT_SITE_URL;
const SHAREPOINT_CLIENT_ID = process.env.SHAREPOINT_CLIENT_ID;
const SHAREPOINT_CLIENT_SECRET = process.env.SHAREPOINT_CLIENT_SECRET;
const SHAREPOINT_TENANT_ID = process.env.SHAREPOINT_TENANT_ID;
const SHAREPOINT_DOCUMENT_LIBRARY = process.env.SHAREPOINT_DOCUMENT_LIBRARY || 'Documents';

// Get access token for SharePoint API
async function getSharePointAccessToken(context) {
    try {
        logMessage(`🔐 Getting SharePoint access token for tenant: ${SHAREPOINT_TENANT_ID}`, context);
        
        const tokenUrl = `https://login.microsoftonline.com/${SHAREPOINT_TENANT_ID}/oauth2/v2.0/token`;
        logMessage(`🔗 Token URL: ${tokenUrl}`, context);
        
        const params = new URLSearchParams();
        params.append('client_id', SHAREPOINT_CLIENT_ID);
        params.append('client_secret', SHAREPOINT_CLIENT_SECRET);
        params.append('scope', `${SHAREPOINT_SITE_URL}/.default`);
        params.append('grant_type', 'client_credentials');
        
        logMessage(`📋 Requesting token with scope: ${SHAREPOINT_SITE_URL}/.default`, context);

        const response = await axios.post(tokenUrl, params, {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        });

        logMessage(`✅ Access token obtained successfully`, context);
        return response.data.access_token;
    } catch (error) {
        logMessage(`❌ SharePoint authentication failed: ${error.message}`, context);
        if (error.response) {
            logMessage(`❌ Response status: ${error.response.status}`, context);
            logMessage(`❌ Response data: ${JSON.stringify(error.response.data)}`, context);
        }
        handleError(error, 'SharePoint Authentication', context);
        throw error;
    }
}

// Upload JSON report to SharePoint
async function uploadJsonToSharePoint(jsonData, fileName, folderPath, context) {
    try {
        logMessage(`📤 Starting JSON upload to SharePoint: ${fileName}`, context);
        logMessage(`📁 Target folder: ${folderPath}`, context);
        
        logMessage(`🔐 Getting access token...`, context);
        const accessToken = await getSharePointAccessToken(context);
        logMessage(`✅ Access token received`, context);
        
        logMessage(`📝 Converting JSON data to buffer...`, context);
        const jsonContent = JSON.stringify(jsonData, null, 2);
        const buffer = Buffer.from(jsonContent, 'utf8');
        logMessage(`📊 JSON buffer size: ${buffer.length} bytes`, context);
        
        const uploadUrl = `${SHAREPOINT_SITE_URL}/_api/web/GetFolderByServerRelativeUrl('${SHAREPOINT_DOCUMENT_LIBRARY}/${folderPath}')/Files/Add(url='${fileName}',overwrite=true)`;
        logMessage(`🔗 Upload URL: ${uploadUrl}`, context);
        
        logMessage(`📤 Sending request to SharePoint...`, context);
        const response = await axios.post(uploadUrl, buffer, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json',
                'Content-Length': buffer.length
            },
            timeout: 30000 // 30 second timeout
        });

        logMessage(`✅ JSON uploaded to SharePoint successfully: ${fileName}`, context);
        logMessage(`📊 Response status: ${response.status}`, context);
        return response.data;
    } catch (error) {
        logMessage(`❌ JSON upload failed for: ${fileName}`, context);
        logMessage(`❌ Error message: ${error.message}`, context);
        if (error.response) {
            logMessage(`❌ Response status: ${error.response.status}`, context);
            logMessage(`❌ Response data: ${JSON.stringify(error.response.data)}`, context);
        }
        if (error.code) {
            logMessage(`❌ Error code: ${error.code}`, context);
        }
        handleError(error, 'SharePoint JSON Upload', context);
        throw error;
    }
}

// Upload PDF report to SharePoint
async function uploadPdfToSharePoint(pdfContent, fileName, folderPath, context) {
    try {
        logMessage(`📤 Starting PDF upload to SharePoint: ${fileName}`, context);
        
        logMessage(`🔐 Getting access token for PDF upload...`, context);
        const accessToken = await getSharePointAccessToken(context);
        logMessage(`✅ Access token received for PDF upload`, context);
        
        // Convert HTML to PDF buffer if needed (requires puppeteer)
        let pdfBuffer;
        if (typeof pdfContent === 'string') {
            logMessage(`🔄 Converting HTML to PDF using Puppeteer...`, context);
            // If pdfContent is HTML string, convert to PDF
            const puppeteer = require('puppeteer');
            const browser = await puppeteer.launch({ headless: true });
            const page = await browser.newPage();
            await page.setContent(pdfContent, { waitUntil: 'networkidle0' });
            pdfBuffer = await page.pdf({
                format: 'A4',
                printBackground: true,
                margin: { top: '20mm', bottom: '20mm', left: '10mm', right: '10mm' }
            });
            await browser.close();
            logMessage(`✅ PDF generated, size: ${pdfBuffer.length} bytes`, context);
        } else {
            pdfBuffer = pdfContent;
            logMessage(`📄 Using provided PDF buffer, size: ${pdfBuffer.length} bytes`, context);
        }
        
        const uploadUrl = `${SHAREPOINT_SITE_URL}/_api/web/GetFolderByServerRelativeUrl('${SHAREPOINT_DOCUMENT_LIBRARY}/${folderPath}')/Files/Add(url='${fileName}',overwrite=true)`;
        logMessage(`🔗 PDF Upload URL: ${uploadUrl}`, context);
        
        logMessage(`📤 Sending PDF to SharePoint...`, context);
        const response = await axios.post(uploadUrl, pdfBuffer, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/pdf',
                'Content-Length': pdfBuffer.length
            },
            timeout: 60000 // 60 second timeout for PDF
        });

        logMessage(`✅ PDF uploaded to SharePoint successfully: ${fileName}`, context);
        return response.data;
    } catch (error) {
        logMessage(`❌ PDF upload failed for: ${fileName}`, context);
        logMessage(`❌ Error message: ${error.message}`, context);
        if (error.response) {
            logMessage(`❌ Response status: ${error.response.status}`, context);
            logMessage(`❌ Response data: ${JSON.stringify(error.response.data)}`, context);
        }
        handleError(error, 'SharePoint PDF Upload', context);
        throw error;
    }
}

// Upload original document to SharePoint
async function uploadOriginalDocumentToSharePoint(base64Content, fileName, folderPath, context) {
    try {
        logMessage(`📤 Starting original document upload to SharePoint: ${fileName}`, context);
        
        logMessage(`🔐 Getting access token for original document...`, context);
        const accessToken = await getSharePointAccessToken(context);
        logMessage(`✅ Access token received for original document`, context);
        
        logMessage(`📄 Converting base64 to buffer...`, context);
        const buffer = Buffer.from(base64Content, 'base64');
        logMessage(`📊 Original document buffer size: ${buffer.length} bytes`, context);
        
        const uploadUrl = `${SHAREPOINT_SITE_URL}/_api/web/GetFolderByServerRelativeUrl('${SHAREPOINT_DOCUMENT_LIBRARY}/${folderPath}')/Files/Add(url='${fileName}',overwrite=true)`;
        logMessage(`🔗 Original document Upload URL: ${uploadUrl}`, context);
        
        logMessage(`📤 Sending original document to SharePoint...`, context);
        const response = await axios.post(uploadUrl, buffer, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json;odata=verbose',
                'Content-Length': buffer.length
            },
            timeout: 60000 // 60 second timeout
        });

        logMessage(`✅ Original document uploaded to SharePoint successfully: ${fileName}`, context);
        return response.data;
    } catch (error) {
        logMessage(`❌ Original document upload failed for: ${fileName}`, context);
        logMessage(`❌ Error message: ${error.message}`, context);
        if (error.response) {
            logMessage(`❌ Response status: ${error.response.status}`, context);
            logMessage(`❌ Response data: ${JSON.stringify(error.response.data)}`, context);
        }
        handleError(error, 'SharePoint Original Document Upload', context);
        throw error;
    }
}

// Create SharePoint folder if it doesn't exist
async function ensureSharePointFolder(folderPath, context) {
    try {
        logMessage(`📁 Ensuring SharePoint folder exists: ${folderPath}`, context);
        
        logMessage(`🔐 Getting access token for folder creation...`, context);
        const accessToken = await getSharePointAccessToken(context);
        logMessage(`✅ Access token received for folder creation`, context);
        
        const folderUrl = `${SHAREPOINT_SITE_URL}/_api/web/folders/add('${SHAREPOINT_DOCUMENT_LIBRARY}/${folderPath}')`;
        logMessage(`🔗 Folder creation URL: ${folderUrl}`, context);
        
        logMessage(`📁 Creating folder...`, context);
        await axios.post(folderUrl, {}, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json'
            }
        });

        logMessage(`✅ SharePoint folder created/ensured: ${folderPath}`, context);
    } catch (error) {
        if (error.response && error.response.status === 409) {
            logMessage(`📁 Folder already exists: ${folderPath}`, context);
        } else {
            logMessage(`❌ Folder creation failed: ${error.message}`, context);
            if (error.response) {
                logMessage(`❌ Response status: ${error.response.status}`, context);
                logMessage(`❌ Response data: ${JSON.stringify(error.response.data)}`, context);
            }
            handleError(error, 'SharePoint Folder Creation', context);
        }
    }
}

module.exports = {
    getSharePointAccessToken,
    uploadJsonToSharePoint,
    uploadPdfToSharePoint,
    uploadOriginalDocumentToSharePoint,
    ensureSharePointFolder
};
