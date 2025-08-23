
function logMessage(message, context) {
    if (context && context.log) {
        context.log(message);
    } else {
        console.log(message);
    }
}

function handleError(error, phase, context) {
    const log = context?.log?.error || console.error;
    log(`[ERROR - ${phase}] ${error.message}`);
    if (error.response) {
        log(`[RESPONSE] ${JSON.stringify(error.response.data, null, 2)}`);
    }
    log(`[STACK] ${error.stack}`);
}

module.exports = {
    logMessage,
    handleError
};