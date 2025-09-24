const { logMessage, handleError, convertHeicToJpegIfNeeded} = require('../utils');
const axios = require('axios');
const FormData = require('form-data');


async function prepareGeneralManagementReport(extractedRows, context, base64BinFile, originalFileName) {
    logMessage("🚀 prepareGeneralManagementReport() called", context);
    logMessage('############################# extractedRows');
    logMessage(extractedRows);
    logMessage('#############################');
    logMessage(base64BinFile);
    logMessage('############################# base64BinFile');
    logMessage(originalFileName);
    logMessage('############################# originalFileName');
}
