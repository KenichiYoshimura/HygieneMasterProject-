const { logMessage, handleError, convertHeicToJpegIfNeeded} = require('../utils');
const axios = require('axios');
const FormData = require('form-data');


async function prepareImportantManagementReport(extractedRows, context, base64BinFile, originalFileName) {
    logMessage("🚀 prepareImportantManagementReport() called", context);
    logMessage('############################# extractedRows');
    logMessage(extractedRows);
    logMessage('############################# base64BinFile');
    logMessage(base64BinFile);
    logMessage('############################# originalFileName');
    logMessage(originalFileName);
    logMessage('#############################');
}