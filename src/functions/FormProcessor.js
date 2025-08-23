const { app } = require('@azure/functions');

app.storageBlob('FormProcessor', {
  // Fire for any blob under the 'incoming-emails' container
  path: 'incoming-emails/{name}',              // or 'incoming-emails/{*path}' to include subfolders
  connection: 'hygienemasterstorage_STORAGE',  // app setting name (see below)
  handler: async (blob, context) => {
    const name = context.triggerMetadata.name;
    context.log(`✅ Processed blob "${name}" (${blob.length} bytes)!!!!!`);
  }
});
