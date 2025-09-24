# HygienMaster

A comprehensive Azure Function App for automated document processing and management system integration, specifically designed for hygiene management forms.

## Overview

HygienMaster automatically processes uploaded documents (PDF, images, HEIC files) by:
1. **Document Classification** - Uses Azure Document Intelligence to classify document types
2. **Data Extraction** - Extracts structured data from classified forms
3. **File Conversion** - Converts HEIC files to JPEG for better compatibility
4. **Integration** - Uploads processed data and files to Monday.com dashboards

## Features

### Document Processing
- **Multi-format Support**: PDF, JPG, PNG, HEIC, and other common formats
- **HEIC Conversion**: Automatically converts HEIC files to JPEG before upload
- **Azure Document Intelligence**: Leverages AI for document classification and data extraction
- **Form Processing**: Specialized extractors for different management form types

### Supported Form Types
- **General Management Forms**: Daily tracking with 7 categories across 7 days
- **Important Management Forms**: Critical management tracking

### Monday.com Integration
- **Automated Upload**: Creates items in Monday.com boards with extracted data
- **File Attachment**: Uploads original (or converted) documents as attachments
- **Data Mapping**: Maps form fields to Monday.com board columns
- **Throttling**: Implements rate limiting to respect API limits

## Project Structure

```
src/
├── functions/
│   ├── docIntelligence/
│   │   ├── documentClassifier.js          # Document classification
│   │   ├── generalManagementFormExtractor.js  # General form data extraction
│   │   ├── importantManagementFormExtractor.js # Important form data extraction
│   │   └── ocrTitleDetector.js            # OCR title detection
│   ├── monday/
│   │   ├── generalManagementDashboard.js  # General management Monday.com integration
│   │   └── importantManagementDashboard.js # Important management Monday.com integration
│   ├── FormProcessor.js                   # Main form processing orchestrator
│   └── utils.js                          # Shared utilities (blob operations, HEIC conversion, etc.)
└── index.js                              # Entry point
```

## Environment Variables

```bash
# Azure Document Intelligence
CLASSIFIER_ENDPOINT=your_azure_endpoint
CLASSIFIER_ENDPOINT_AZURE_API_KEY=your_azure_api_key
CLASSIFIER_ID=your_classifier_id

# Monday.com API
MONDAY_API_KEY=your_monday_api_token

# Azure Storage
AZURE_STORAGE_CONNECTION_STRING=your_storage_connection_string
```

## Installation

1. Clone the repository
```bash
git clone https://github.com/your-username/HygienMaster.git
cd HygienMaster
```

2. Install dependencies
```bash
npm install
```

3. Configure environment variables
```bash
cp .env.example .env
# Edit .env with your actual values
```

4. Deploy to Azure Functions
```bash
# Using Azure Functions Core Tools
func azure functionapp publish your-function-app-name
```

## Dependencies

### Core Dependencies
- `@azure/storage-blob` - Azure Blob Storage operations
- `axios` - HTTP requests for API calls
- `form-data` - Multipart form data for file uploads
- `heic-convert` - HEIC to JPEG conversion
- `mime` - MIME type detection

### Development Dependencies
- `dotenv` - Environment variable management (local development)

## Usage

### Document Processing Flow

1. **Upload**: Documents are uploaded to Azure Blob Storage
2. **Trigger**: Blob trigger activates the function
3. **Classification**: Document is classified using Azure Document Intelligence
4. **Extraction**: Relevant data is extracted based on document type
5. **Conversion**: HEIC files are converted to JPEG if needed
6. **Upload**: Data and files are uploaded to Monday.com
7. **Archive**: Processed documents are moved to appropriate storage containers

### Monday.com Board Structure

#### General Management Board
- **Columns**: Date, Store, 7 Categories, Comments, Approver, File attachment
- **Data Mapping**: Extracts daily tracking data across multiple categories

#### Important Management Board  
- **Columns**: Date, Store, Status indicators, File attachment
- **Data Mapping**: Focuses on critical management items

## API Integration

### Azure Document Intelligence
- **Classification**: Automatic document type detection
- **Field Extraction**: Structured data extraction from forms
- **OCR**: Text recognition from scanned documents

### Monday.com API
- **GraphQL**: Uses Monday.com's GraphQL API for data operations
- **File Upload**: Handles file attachments via multipart uploads
- **Rate Limiting**: Implements proper throttling for API stability

## Error Handling

- Comprehensive logging for troubleshooting
- Graceful error handling with detailed error messages
- Automatic retry mechanisms for transient failures
- File movement to error containers for failed processing

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For support and questions, please open an issue in the GitHub repository.
