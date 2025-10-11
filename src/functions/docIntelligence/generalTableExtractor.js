'use strict';

if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}

const { AzureKeyCredential, DocumentAnalysisClient } = require("@azure/ai-form-recognizer");
const { logMessage, handleError } = require('../utils');
const axios = require('axios');

/* -----------------------------------------------------------------------------
  Azure setup
----------------------------------------------------------------------------- */
const endpoint = process.env['CLASSIFIER_ENDPOINT'];
const apiKey = process.env['CLASSIFIER_ENDPOINT_AZURE_API_KEY'];

// ✅ Custom Vision setup using environment variables
const customVisionEndpoint = process.env['CUSTOM_VISION_ENDPOINT'] || 'https://customsymbolrecognizer.cognitiveservices.azure.com/';
const customVisionKey = process.env['CUSTOM_VISION_KEY'];
const customVisionProjectId = process.env['CUSTOM_VISION_PROJECT_ID'];
const customVisionIterationName = process.env['CUSTOM_VISION_ITERATION_NAME'] || 'Iteration1';

logMessage('🔧 Document Intelligence Endpoint: ' + endpoint, null);
logMessage('🔧 Document Intelligence API Key: ' + (apiKey ? '[REDACTED]' : '❌ Missing API Key'), null);
logMessage('🔧 Custom Vision Endpoint: ' + customVisionEndpoint, null);
logMessage('🔧 Custom Vision API Key: ' + (customVisionKey ? '[REDACTED]' : '❌ Missing API Key'), null);

if (!endpoint || !apiKey) {
  throw new Error('Missing CLASSIFIER_ENDPOINT or CLASSIFIER_ENDPOINT_AZURE_API_KEY');
}
if (!customVisionEndpoint || !customVisionKey) {
  throw new Error('Missing CUSTOM_VISION_ENDPOINT or CUSTOM_VISION_KEY');
}

const client = new DocumentAnalysisClient(endpoint, new AzureKeyCredential(apiKey));

/* -----------------------------------------------------------------------------
  OCR and Custom Symbol Recognition Functions
----------------------------------------------------------------------------- */

/**
 * Run OCR on a specific cell region of the image
 */
async function runCellOCR(imageBuffer, bbox, context) {
  try {
    if (!bbox) {
      logMessage('⚠️ No bounding box provided for OCR', context);
      return { success: false, text: '', confidence: 0 };
    }

    logMessage(`🔍 Running OCR on cell region: [${bbox.join(', ')}]`, context);
    
    const ocrPoller = await client.beginAnalyzeDocument("prebuilt-read", imageBuffer);
    const ocrResult = await ocrPoller.pollUntilDone();
    
    // Find text that intersects with the cell bbox
    let cellText = '';
    let totalConfidence = 0;
    let textCount = 0;

    if (ocrResult.pages) {
      for (const page of ocrResult.pages) {
        if (page.lines) {
          for (const line of page.lines) {
            // Check if line intersects with cell bbox
            if (line.polygon && isTextInBbox(line.polygon, bbox)) {
              cellText += (cellText ? ' ' : '') + line.content;
              if (line.confidence) {
                totalConfidence += line.confidence;
                textCount++;
              }
            }
          }
        }
      }
    }

    const avgConfidence = textCount > 0 ? totalConfidence / textCount : 0;
    
    logMessage(`✅ OCR result: "${cellText}" (confidence: ${Math.round(avgConfidence * 100)}%)`, context);
    
    return {
      success: true,
      text: cellText.trim(),
      confidence: avgConfidence
    };

  } catch (error) {
    logMessage(`❌ OCR failed: ${error.message}`, context);
    return { success: false, text: '', confidence: 0, error: error.message };
  }
}

/**
 * Run custom symbol recognition on a specific cell region
 */
async function runCustomSymbolRecognition(imageBuffer, bbox, context) {
  try {
    if (!bbox) {
      logMessage('⚠️ No bounding box provided for symbol recognition', context);
      return { success: false, predictions: [], confidence: 0 };
    }

    logMessage(`🎯 Running custom symbol recognition on cell region: [${bbox.join(', ')}]`, context);

    // ✅ Use environment variables for Custom Vision API
    const customVisionUrl = `${customVisionEndpoint}customvision/v3.0/Prediction/${customVisionProjectId}/classify/iterations/${customVisionIterationName}/image`;
    
    const response = await axios.post(customVisionUrl, 
      imageBuffer,
      {
        headers: {
          'Prediction-Key': customVisionKey,
          'Content-Type': 'application/octet-stream'
        }
      }
    );

    const predictions = response.data.predictions || [];
    
    logMessage(`✅ Symbol recognition API response received`, context);
    
    // Log all predictions in the format you specified
    if (predictions.length > 0) {
      logMessage(`📊 Symbol Recognition Results:`, context);
      logMessage(`Tag\t\tProbability`, context);
      predictions.forEach((pred) => {
        const probability = (pred.probability * 100).toFixed(1);
        logMessage(`${pred.tagName}\t\t${probability}%`, context);
      });
    }

    const maxConfidence = predictions.length > 0 
      ? Math.max(...predictions.map(p => p.probability || 0))
      : 0;

    const topPrediction = predictions.length > 0 
      ? predictions.reduce((prev, current) => 
          (prev.probability > current.probability) ? prev : current
        )
      : null;

    return {
      success: true,
      predictions: predictions.map(pred => ({
        tagName: pred.tagName,
        probability: pred.probability,
        percentage: `${(pred.probability * 100).toFixed(1)}%`
      })),
      confidence: maxConfidence,
      symbolCount: predictions.length,
      topPrediction: topPrediction
    };

  } catch (error) {
    logMessage(`❌ Custom symbol recognition failed: ${error.message}`, context);
    if (error.response) {
      logMessage(`❌ API Error Status: ${error.response.status}`, context);
      logMessage(`❌ API Error Data: ${JSON.stringify(error.response.data)}`, context);
    }
    return { success: false, predictions: [], confidence: 0, error: error.message };
  }
}

/**
 * Helper function to check if text polygon intersects with bbox
 */
function isTextInBbox(polygon, bbox) {
  if (!polygon || polygon.length < 8 || !bbox || bbox.length < 4) {
    return false;
  }

  const xCoords = polygon.filter((_, i) => i % 2 === 0);
  const yCoords = polygon.filter((_, i) => i % 2 === 1);
  const polyBbox = [
    Math.min(...xCoords),
    Math.min(...yCoords),
    Math.max(...xCoords),
    Math.max(...yCoords)
  ];

  return isBboxOverlapping(polyBbox, bbox);
}

/**
 * Helper function to check if two bounding boxes overlap
 */
function isBboxOverlapping(bbox1, bbox2) {
  const [x1, y1, x2, y2] = bbox1;
  const [x3, y3, x4, y4] = bbox2;
  
  return !(x2 < x3 || x4 < x1 || y2 < y3 || y4 < y1);
}

/* -----------------------------------------------------------------------------
  Analysis pipeline
----------------------------------------------------------------------------- */
async function analyseAndExtract(buffer, mimeType, context) {
  try {
    logMessage('Starting analyseAndExtract...', context);
    console.time('analyseAndExtract');

    logMessage('Calling Azure prebuilt-layout...', context);
    const layoutPoller = await client.beginAnalyzeDocument("prebuilt-layout", buffer, { contentType: mimeType });
    const layoutResult = await layoutPoller.pollUntilDone();

    logMessage('✅ Layout analysis completed', context);

    // Check for table structures
    if (!layoutResult.tables || layoutResult.tables.length === 0) {
      logMessage('❌ No table structures detected in the document', context);
      return [];
    }

    logMessage(`📊 Found ${layoutResult.tables.length} table(s) in the document`, context);

    const merged = [];
    
    // Process each table
    for (let tableIndex = 0; tableIndex < layoutResult.tables.length; tableIndex++) {
      const table = layoutResult.tables[tableIndex];
      logMessage(`🔍 Processing table ${tableIndex + 1}: ${table.rowCount} rows × ${table.columnCount} columns`, context);

      // Create a 2D array to organize cells by position
      const cellGrid = Array(table.rowCount).fill(null).map(() => Array(table.columnCount).fill(null));
      
      // Fill the grid with cells
      for (const cell of table.cells) {
        if (cell.rowIndex < table.rowCount && cell.columnIndex < table.columnCount) {
          cellGrid[cell.rowIndex][cell.columnIndex] = cell;
        }
      }

      // Iterate through cells from top-left, moving right then down
      logMessage(`📋 Iterating through table ${tableIndex + 1} cells:`, context);
      
      for (let row = 0; row < table.rowCount; row++) {
        for (let col = 0; col < table.columnCount; col++) {
          const cell = cellGrid[row][col];
          
          if (cell) {
            const coordinate = `(Row: ${row}, Col: ${col})`;
            const content = cell.content || '';
            const boundingRegions = cell.boundingRegions || [];
            
            logMessage(`  📍 ${coordinate}: "${content}"`, context);
            
            // Extract bounding box if available
            let bbox = null;
            if (boundingRegions.length > 0 && boundingRegions[0].polygon) {
              const polygon = boundingRegions[0].polygon;
              const xCoords = polygon.filter((_, i) => i % 2 === 0);
              const yCoords = polygon.filter((_, i) => i % 2 === 1);
              bbox = [
                Math.min(...xCoords),
                Math.min(...yCoords),
                Math.max(...xCoords),
                Math.max(...yCoords)
              ];
            }

            // Run OCR on this cell
            logMessage(`    🔍 Running OCR for cell ${coordinate}...`, context);
            const ocrResult = await runCellOCR(buffer, bbox, context);

            // Run custom symbol recognition on this cell
            logMessage(`    🎯 Running symbol recognition for cell ${coordinate}...`, context);
            const symbolResult = await runCustomSymbolRecognition(buffer, bbox, context);

            // Add cell to merged results
            merged.push({
              tableIndex: tableIndex,
              rowIndex: row,
              columnIndex: col,
              coordinate: coordinate,
              displayText: content,
              content: content,
              bbox: bbox,
              boundingRegions: boundingRegions,
              rowSpan: cell.rowSpan || 1,
              columnSpan: cell.columnSpan || 1,
              kind: cell.kind || 'content',
              isHandwritten: false,
              orientationDeg: 0,
              confidence: cell.confidence || 0,
              ocrResult: {
                success: ocrResult.success,
                text: ocrResult.text || '',
                confidence: ocrResult.confidence || 0,
                error: ocrResult.error || null
              },
              symbolResult: {
                success: symbolResult.success,
                predictions: symbolResult.predictions || [],
                confidence: symbolResult.confidence || 0,
                symbolCount: symbolResult.symbolCount || 0,
                error: symbolResult.error || null
              },
              tableMetadata: {
                tableIndex: tableIndex,
                totalRows: table.rowCount,
                totalColumns: table.columnCount,
                tableCaption: table.caption || null,
                tableBoundingRegions: table.boundingRegions || []
              }
            });

            // Log results summary
            if (ocrResult.success || symbolResult.success) {
              logMessage(`    ✅ Cell ${coordinate} analysis complete:`, context);
              if (ocrResult.success) {
                logMessage(`       📝 OCR: "${ocrResult.text}" (${Math.round(ocrResult.confidence * 100)}%)`, context);
              }
              if (symbolResult.success && symbolResult.symbolCount > 0) {
                logMessage(`       🎯 Symbols: ${symbolResult.symbolCount} detected (${Math.round(symbolResult.confidence * 100)}%)`, context);
              }
            }

          } else {
            // Empty cell
            const coordinate = `(Row: ${row}, Col: ${col})`;
            logMessage(`  📍 ${coordinate}: [EMPTY CELL]`, context);
            
            merged.push({
              tableIndex: tableIndex,
              rowIndex: row,
              columnIndex: col,
              coordinate: coordinate,
              displayText: '',
              content: '',
              bbox: null,
              boundingRegions: [],
              rowSpan: 1,
              columnSpan: 1,
              kind: 'empty',
              isHandwritten: false,
              orientationDeg: 0,
              confidence: 0,
              ocrResult: { success: false, text: '', confidence: 0 },
              symbolResult: { success: false, predictions: [], confidence: 0, symbolCount: 0 },
              tableMetadata: {
                tableIndex: tableIndex,
                totalRows: table.rowCount,
                totalColumns: table.columnCount,
                tableCaption: table.caption || null,
                tableBoundingRegions: table.boundingRegions || []
              }
            });
          }
        }
      }
      
      logMessage(`✅ Completed processing table ${tableIndex + 1}`, context);
    }

    logMessage(`📊 Table analysis complete! Extracted ${merged.length} cells from ${layoutResult.tables.length} table(s)`, context);
    
    const ocrSuccessCount = merged.filter(cell => cell.ocrResult.success).length;
    const symbolSuccessCount = merged.filter(cell => cell.symbolResult.success && cell.symbolResult.symbolCount > 0).length;
    
    logMessage(`📈 Analysis Summary:`, context);
    logMessage(`   📝 OCR successful: ${ocrSuccessCount}/${merged.length} cells`, context);
    logMessage(`   🎯 Symbols detected: ${symbolSuccessCount}/${merged.length} cells`, context);
    
    console.timeEnd('analyseAndExtract');
    
    return merged;

  } catch (error) {
    handleError(error, 'analyseAndExtract', context);
    return null;
  }
}

/* -----------------------------------------------------------------------------
  Complete processing pipeline for table documents
----------------------------------------------------------------------------- */
async function processUnknownDocumentWithTables(imageBuffer, mimeType, base64Raw, originalFileName, companyName, context) {
  try {
    logMessage(`🧠 Starting complete table document processing pipeline`, context);
    logMessage(`📄 File: ${originalFileName}, Company: ${companyName}, MIME: ${mimeType}`, context);

    // Step 1: Extract and analyze the document
    logMessage(`📖 Starting document analysis...`, context);
    const analyseOutput = await analyseAndExtract(imageBuffer, mimeType, context);

    if (!analyseOutput || analyseOutput.length === 0) {
      logMessage(`❌ No table structures detected in the document`, context);
      return {
        success: false,
        reason: 'no_tables_detected',
        textRegions: 0,
        uploads: { original: false, json: false, annotatedImage: false }
      };
    }

    logMessage(`✅ Analysis complete! Found ${analyseOutput.length} table cells`, context);
    
    // Log sample extracted data for debugging
    if (analyseOutput.length > 0) {
      logMessage(`📝 Sample extracted table cells:`, context);
      analyseOutput.slice(0, 3).forEach((cell, idx) => {
        const text = cell.displayText || '';
        const coordinate = cell.coordinate || '';
        const ocrText = cell.ocrResult?.text || '';
        const symbolCount = cell.symbolResult?.symbolCount || 0;
        logMessage(`  [${idx}] ${coordinate} Table: "${text}" | OCR: "${ocrText}" | Symbols: ${symbolCount}`, context);
      });
    }

    // Prepare SharePoint variables
    const baseName = originalFileName.replace(/\.[^/.]+$/, "");
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    
    // JST time for folder
    const currentDateUTC = new Date();
    const currentDateJST = new Date(currentDateUTC.getTime() + (9 * 60 * 60 * 1000));
    const dateFolder = [
      currentDateJST.getUTCFullYear(),
      String(currentDateJST.getUTCMonth() + 1).padStart(2, '0'),
      String(currentDateJST.getUTCDate()).padStart(2, '0')
    ].join('-') + '_' + [
      String(currentDateJST.getUTCHours()).padStart(2, '0'),
      String(currentDateJST.getUTCMinutes()).padStart(2, '0'),
      String(currentDateJST.getUTCSeconds()).padStart(2, '0'),
      String(currentDateJST.getUTCMilliseconds()).padStart(3, '0')
    ].join('-');
    
    const basePath = process.env.SHAREPOINT_ETC_FOLDER_PATH?.replace(/^\/+|\/+$/g, '') || 'その他';
    const folderPath = `${basePath}/${companyName}/${dateFolder}`;
    logMessage(`📁 Target SharePoint folder (JST): ${folderPath}`, context);

    // Import SharePoint functions
    const { ensureSharePointFolder, uploadJsonToSharePoint, uploadOriginalDocumentToSharePoint } = require('../sharepoint/sendToSharePoint');
    
    // Ensure folder exists
    await ensureSharePointFolder(folderPath, context);

    const uploadResults = { original: false, json: false, annotatedImage: false };

    // Step 2: Upload original document
    logMessage(`📤 Uploading original document to SharePoint...`, context);
    const originalDocFileName = `original-${originalFileName}`;
    
    try {
      const originalDocUploadResult = await uploadOriginalDocumentToSharePoint(
        base64Raw, originalDocFileName, folderPath, context
      );
      
      if (originalDocUploadResult) {
        logMessage(`✅ Successfully uploaded original document: ${originalDocFileName}`, context);
        uploadResults.original = true;
      } else {
        logMessage(`⚠️ Failed to upload original document, but continuing...`, context);
      }
    } catch (error) {
      logMessage(`❌ Error uploading original document: ${error.message}`, context);
    }

    // Step 3: Upload JSON analysis
    logMessage(`📤 Uploading table analysis JSON to SharePoint...`, context);
    
    const analysisJsonReport = {
      metadata: {
        originalFileName, 
        processedDate: new Date().toISOString(),
        companyName, 
        mimeType,
        totalTableCells: analyseOutput.length,
        tablesDetected: analyseOutput.length > 0 ? Math.max(...analyseOutput.map(cell => cell.tableIndex)) + 1 : 0,
        ocrSuccessCount: analyseOutput.filter(cell => cell.ocrResult.success).length,
        symbolDetectionCount: analyseOutput.filter(cell => cell.symbolResult.success && cell.symbolResult.symbolCount > 0).length,
        documentType: 'table_extraction_with_symbols'
      },
      extractedData: {
        tableCells: analyseOutput,
        extractionMethod: 'Azure Document Intelligence (Layout) + Custom OCR + Custom Symbol Recognition',
        ocrEnabled: true,
        symbolRecognitionEnabled: true,
        customVisionEndpoint: customVisionEndpoint
      },
      summary: {
        tableStructures: analyseOutput.length > 0 ? Math.max(...analyseOutput.map(cell => cell.tableIndex)) + 1 : 0,
        totalCells: analyseOutput.length,
        cellsWithText: analyseOutput.filter(cell => cell.ocrResult.text).length,
        cellsWithSymbols: analyseOutput.filter(cell => cell.symbolResult.symbolCount > 0).length
      }
    };
    
    const jsonFileName = `テーブル抽出データ-${baseName}-${timestamp}.json`;
    
    try {
      const jsonUploadResult = await uploadJsonToSharePoint(
        analysisJsonReport, jsonFileName, folderPath, context
      );
      
      if (jsonUploadResult) {
        logMessage(`✅ Successfully uploaded table analysis JSON: ${jsonFileName}`, context);
        uploadResults.json = true;
      } else {
        logMessage(`⚠️ Failed to upload analysis JSON, but continuing...`, context);
      }
    } catch (error) {
      logMessage(`❌ Error uploading analysis JSON: ${error.message}`, context);
    }

    // Return comprehensive results
    const successfulUploads = Object.values(uploadResults).filter(Boolean).length;
    const totalUploads = Object.keys(uploadResults).length;
    
    logMessage(`📊 Processing complete! Table cells: ${analyseOutput.length}, Uploads: ${successfulUploads}/${totalUploads}`, context);
    
    return {
      success: true,
      textRegions: analyseOutput.length,
      tablesDetected: analyseOutput.length > 0 ? Math.max(...analyseOutput.map(cell => cell.tableIndex)) + 1 : 0,
      ocrSuccessCount: analyseOutput.filter(cell => cell.ocrResult.success).length,
      symbolDetectionCount: analyseOutput.filter(cell => cell.symbolResult.success && cell.symbolResult.symbolCount > 0).length,
      uploads: uploadResults,
      sharePointFolder: folderPath,
      analysisData: analyseOutput
    };

  } catch (error) {
    handleError(error, 'processUnknownDocument', context);
    return {
      success: false,
      reason: 'processing_error',
      error: error.message,
      textRegions: 0,
      uploads: { original: false, json: false, annotatedImage: false }
    };
  }
}

/* -----------------------------------------------------------------------------
  Exports (CommonJS)
----------------------------------------------------------------------------- */
module.exports = {
  analyseAndExtract,
  runCellOCR,
  runCustomSymbolRecognition,
  processUnknownDocumentWithTables
};