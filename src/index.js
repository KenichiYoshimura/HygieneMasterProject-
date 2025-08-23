const { app } = require('@azure/functions');

app.setup({
    enableHttpStream: true,
});


require('./functions/FormProcessor');