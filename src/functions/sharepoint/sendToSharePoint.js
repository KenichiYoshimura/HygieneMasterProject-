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
        const tokenUrl = `https://login.microsoftonline.com/${SHAREPOINT_TENANT_ID}/oauth2/v2.0/token`;
        
        const params = new URLSearchParams();
        params.append('client_id', SHAREPOINT_CLIENT_ID);
        params.append('client_secret', SHAREPOINT_CLIENT_SECRET);
        params.append('scope', `${SHAREPOINT_SITE_URL}/.default`);
        params.append('grant_type', 'client_credentials');

        const response = await axios.post(tokenUrl, params, {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        });

        return response.data.access_token;
    } catch (error) {
        handleError(error, 'SharePoint Authentication', context);
        throw error;
    }
}

// Upload JSON report to SharePoint
async function uploadJsonToSharePoint(jsonData, fileName, folderPath, context) {
    try {
        logMessage(`📤 Uploading JSON to SharePoint: ${fileName}`, context);
        
        const accessToken = await getSharePointAccessToken(context);
        const jsonContent = JSON.stringify(jsonData, null, 2);
        const buffer = Buffer.from(jsonContent, 'utf8');
        
        const uploadUrl = `${SHAREPOINT_SITE_URL}/_api/web/GetFolderByServerRelativeUrl('${SHAREPOINT_DOCUMENT_LIBRARY}/${folderPath}')/Files/Add(url='${fileName}',overwrite=true)`;
        
        const response = await axios.post(uploadUrl, buffer, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json',
                'Content-Length': buffer.length
            }
        });

        logMessage(`✅ JSON uploaded to SharePoint: ${fileName}`, context);
        return response.data;
    } catch (error) {
        handleError(error, 'SharePoint JSON Upload', context);
        throw error;
    }
}

// Upload PDF report to SharePoint
async function uploadPdfToSharePoint(pdfContent, fileName, folderPath, context) {
    try {
        logMessage(`📤 Uploading PDF to SharePoint: ${fileName}`, context);
        
        const accessToken = await getSharePointAccessToken(context);
        
        // Convert HTML to PDF buffer if needed (requires puppeteer)
        let pdfBuffer;
        if (typeof pdfContent === 'string') {
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
        } else {
            pdfBuffer = pdfContent;
        }
        
        const uploadUrl = `${SHAREPOINT_SITE_URL}/_api/web/GetFolderByServerRelativeUrl('${SHAREPOINT_DOCUMENT_LIBRARY}/${folderPath}')/Files/Add(url='${fileName}',overwrite=true)`;
        
        const response = await axios.post(uploadUrl, pdfBuffer, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/pdf',
                'Content-Length': pdfBuffer.length
            }
        });

        logMessage(`✅ PDF uploaded to SharePoint: ${fileName}`, context);
        return response.data;
    } catch (error) {
        handleError(error, 'SharePoint PDF Upload', context);
        throw error;
    }
}

// Upload original document to SharePoint
async function uploadOriginalDocumentToSharePoint(base64Content, fileName, folderPath, context) {
    try {
        logMessage(`📤 Uploading original document to SharePoint: ${fileName}`, context);
        
        const accessToken = await getSharePointAccessToken(context);
        const buffer = Buffer.from(base64Content, 'base64');
        
        const uploadUrl = `${SHAREPOINT_SITE_URL}/_api/web/GetFolderByServerRelativeUrl('${SHAREPOINT_DOCUMENT_LIBRARY}/${folderPath}')/Files/Add(url='${fileName}',overwrite=true)`;
        
        const response = await axios.post(uploadUrl, buffer, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json;odata=verbose',
                'Content-Length': buffer.length
            }
        });

        logMessage(`✅ Original document uploaded to SharePoint: ${fileName}`, context);
        return response.data;
    } catch (error) {
        handleError(error, 'SharePoint Original Document Upload', context);
        throw error;
    }
}

// Create SharePoint folder if it doesn't exist
async function ensureSharePointFolder(folderPath, context) {
    try {
        const accessToken = await getSharePointAccessToken(context);
        
        const folderUrl = `${SHAREPOINT_SITE_URL}/_api/web/folders/add('${SHAREPOINT_DOCUMENT_LIBRARY}/${folderPath}')`;
        
        await axios.post(folderUrl, {}, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json'
            }
        });

        logMessage(`📁 SharePoint folder ensured: ${folderPath}`, context);
    } catch (error) {
        // Folder might already exist, which is fine
        if (error.response && error.response.status !== 409) {
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
