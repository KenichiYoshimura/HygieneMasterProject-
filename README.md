
# HygienMaster

HygienMaster is an Azure Functions-based document processing system designed to classify and extract structured data from uploaded hygiene management forms, and upload the results to a Monday.com dashboard.

## 📁 Project Structure

```
HygienMaster/
├── .env                        # Environment variables (local only)
├── host.json                   # Azure Functions host configuration
├── local.settings.json         # Local development settings
├── package.json                # Node.js dependencies
├── src/
│   ├── index.js                # Entry point (optional)
│   └── functions/
│       ├── FormProcessor.js    # Azure Function triggered by blob upload
│       ├── utils.js            # Logging and error handling utilities
│       ├── monday/
│       │   └── importantManagementDashboard.js  # Upload logic to Monday.com
│       └── docIntelligence/
│           ├── documentClassifier.js            # Classifies uploaded documents
│           └── importantManagementFormExtractor.js # Extracts structured data
```

## ⚙️ Setup Instructions

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/HygienMaster.git
   cd HygienMaster
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Create a `.env` file with the following variables:
   ```env
   EXTRACTOR_ENDPOINT=...
   EXTRACTOR_ENDPOINT_AZURE_API_KEY=...
   EXTRACTOR_MODEL_ID=...
   MONDAY_API_KEY=...
   ```

4. Ensure `local.settings.json` includes:
   ```json
   {
     "IsEncrypted": false,
     "Values": {
       "AzureWebJobsStorage": "UseDevelopmentStorage=true",
       "FUNCTIONS_WORKER_RUNTIME": "node"
     }
   }
   ```

## 🚀 Running Locally

1. Start the Azure Functions runtime:
   ```bash
   func start
   ```

2. Upload a file to the `incoming-emails/` blob container to trigger the `FormProcessor` function.

## 📦 Deployment

To deploy to Azure:
```bash
func azure functionapp publish <your-function-app-name>
```

---

## 🧪 Testing
- Place test files in the `incoming-emails/` container.
- Monitor logs via `func start` or Azure Log Stream.

## 📄 License
MIT License
