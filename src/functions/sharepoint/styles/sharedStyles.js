/**
 * Shared CSS styles for hygiene management reports
 * @param {string} theme - 'general' or 'important' for theme-specific colors
 * @returns {string} CSS styles as string
 */
function getReportStyles(theme = 'general') {
    const themeColors = {
        general: {
            headerGradient: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
            thColor: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
            borderColor: '#3498db',
            headerIcon: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="white" opacity="0.1"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/></svg>'
        },
        important: {
            headerGradient: 'linear-gradient(135deg, #2e7d32 0%, #388e3c 100%)',
            thColor: 'linear-gradient(135deg, #2e7d32 0%, #388e3c 100%)',
            borderColor: '#4caf50',
            headerIcon: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="white" opacity="0.1"><path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"/></svg>'
        }
    };

    const colors = themeColors[theme] || themeColors.general;

    return `
        /* Modern Professional Styling */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', 'Yu Gothic', 'Meiryo', sans-serif;
            line-height: 1.6;
            color: #2c3e50;
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            min-height: 100vh;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background: #ffffff;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            border-radius: 12px;
            margin-top: 20px;
            margin-bottom: 20px;
        }

        /* Header Section */
        .header {
            background: ${colors.headerGradient};
            color: white;
            padding: 30px;
            border-radius: 12px 12px 0 0;
            margin: -20px -20px 30px -20px;
            position: relative;
            overflow: hidden;
        }

        .header::before {
            content: '';
            position: absolute;
            top: 0;
            right: 0;
            width: 100px;
            height: 100px;
            background: url('data:image/svg+xml,${colors.headerIcon}') no-repeat center;
        }

        .header h1 {
            font-size: 2.2em;
            font-weight: 300;
            margin-bottom: 10px;
        }

        .header .subtitle {
            font-size: 1.1em;
            opacity: 0.9;
        }

        /* Executive Summary Cards */
        .summary-cards {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }

        .summary-card {
            background: white;
            padding: 25px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            border-left: 4px solid ${colors.borderColor};
            transition: transform 0.2s ease;
        }

        .summary-card:hover {
            transform: translateY(-2px);
        }

        .summary-card.compliance {
            border-left-color: var(--compliance-color);
        }

        .summary-card.daily-check {
            border-left-color: var(--daily-check-color);
        }

        .summary-card.comments {
            border-left-color: #9b59b6;
        }

        .summary-card.sentiment {
            border-left-color: #1abc9c;
        }

        .card-header {
            display: flex;
            align-items: center;
            margin-bottom: 15px;
        }

        .card-icon {
            font-size: 2em;
            margin-right: 15px;
        }

        .card-title {
            font-size: 1.1em;
            font-weight: 600;
            color: #34495e;
        }

        .card-value {
            font-size: 2.5em;
            font-weight: 300;
            color: #2c3e50;
            margin-bottom: 10px;
        }

        .card-description {
            color: #7f8c8d;
            font-size: 0.95em;
        }

        /* Progress Bar */
        .progress-bar {
            width: 100%;
            height: 8px;
            background: #ecf0f1;
            border-radius: 4px;
            overflow: hidden;
            margin-top: 15px;
        }

        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, #27ae60, #2ecc71);
            border-radius: 4px;
            transition: width 0.3s ease;
        }

        .progress-fill.warning {
            background: linear-gradient(90deg, #f39c12, #f1c40f);
        }

        .progress-fill.danger {
            background: linear-gradient(90deg, #e74c3c, #ec7063);
        }

        /* Section Styling */
        .section {
            background: white;
            margin-bottom: 30px;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        }

        .section-header {
            background: #f8f9fa;
            padding: 20px 30px;
            border-bottom: 1px solid #dee2e6;
            cursor: pointer;
            transition: background-color 0.2s ease;
        }

        .section-header:hover {
            background: #e9ecef;
        }

        .section-header h3 {
            color: #2c3e50;
            font-size: 1.3em;
            font-weight: 600;
            display: flex;
            align-items: center;
        }

        .section-header .toggle-icon {
            margin-left: auto;
            transition: transform 0.3s ease;
        }

        .section-content {
            padding: 30px;
        }

        /* Table Styling */
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
            background: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }

        th {
            background: ${colors.thColor};
            color: white;
            padding: 15px 10px;
            font-weight: 600;
            text-align: center;
            font-size: 0.9em;
            letter-spacing: 0.5px;
        }

        td {
            padding: 12px 10px;
            text-align: center;
            border-bottom: 1px solid #f1f3f4;
            vertical-align: top;
        }

        .data-row:hover {
            background: #f8f9fa;
            transition: background-color 0.2s ease;
        }

        .date-cell {
            font-weight: 600;
            background: #f8f9fa;
            color: #2c3e50;
        }

        .comment-cell, .comment-text {
            text-align: left;
            max-width: 200px;
            font-size: 0.9em;
            color: #495057;
        }

        .translation-text {
            text-align: left;
            max-width: 150px;
            font-size: 0.9em;
            color: #495057;
        }

        .language-tag {
            text-align: center;
        }

        .confidence-cell {
            min-width: 120px;
            vertical-align: top;
        }

        /* Status Badges */
        .status-badge {
            display: inline-block;
            padding: 6px 12px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        .status-good {
            background: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }

        .status-bad {
            background: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }

        .status-none {
            background: #e2e3e5;
            color: #6c757d;
            border: 1px solid #d6d8db;
        }

        .status-neutral {
            background: #fff3cd;
            color: #856404;
            border: 1px solid #ffeaa7;
        }

        /* Sentiment Styling */
        .sentiment-badge {
            display: inline-flex;
            align-items: center;
            padding: 6px 12px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: 600;
        }

        .sentiment-positive {
            background: #d4edda;
            color: #155724;
        }

        .sentiment-negative {
            background: #f8d7da;
            color: #721c24;
        }

        .sentiment-neutral {
            background: #e2e3e5;
            color: #6c757d;
        }

        /* Enhanced Language and Sentiment Display Styles */
        .language-badge {
            background: linear-gradient(45deg, #3498db, #2980b9);
            color: white;
            padding: 6px 12px;
            border-radius: 15px;
            font-size: 0.85em;
            font-weight: 600;
            text-shadow: 0 1px 2px rgba(0,0,0,0.1);
            box-shadow: 0 2px 4px rgba(52, 152, 219, 0.3);
            display: inline-block;
            min-width: 60px;
            text-align: center;
            transition: all 0.2s ease;
        }

        .language-badge:hover {
            transform: translateY(-1px);
            box-shadow: 0 4px 8px rgba(52, 152, 219, 0.4);
        }

        .no-translation {
            background: linear-gradient(45deg, #2e7d32, #388e3c);
            color: white;
            padding: 6px 12px;
            border-radius: 15px;
            font-size: 0.85em;
            font-weight: 600;
            text-shadow: 0 1px 2px rgba(0,0,0,0.1);
            box-shadow: 0 2px 4px rgba(46, 125, 50, 0.3);
            display: inline-block;
        }

        /* Enhanced Confidence Display */
        .confidence-container {
            position: relative;
            background: #f0f0f0;
            border-radius: 12px;
            height: 24px;
            overflow: hidden;
            margin-bottom: 8px;
            box-shadow: inset 0 2px 4px rgba(0,0,0,0.1);
        }

        .confidence-fill {
            height: 100%;
            border-radius: 12px;
            transition: width 0.8s ease;
            position: relative;
            overflow: hidden;
        }

        .confidence-fill.sentiment-positive {
            background: linear-gradient(90deg, #27ae60, #2ecc71);
        }

        .confidence-fill.sentiment-neutral {
            background: linear-gradient(90deg, #f39c12, #f1c40f);
        }

        .confidence-fill.sentiment-negative {
            background: linear-gradient(90deg, #e74c3c, #ec7063);
        }

        .confidence-fill::before {
            content: '';
            position: absolute;
            top: 0;
            left: -100%;
            width: 100%;
            height: 100%;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.4), transparent);
            animation: shimmer 2s infinite;
        }

        @keyframes shimmer {
            0% { left: -100%; }
            100% { left: 100%; }
        }

        .confidence-text {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            font-weight: 600;
            font-size: 0.8em;
            color: #2c3e50;
            text-shadow: 0 1px 2px rgba(255,255,255,0.8);
            z-index: 2;
        }

        /* Simplified Inline Confidence Details */
        .inline-confidence-details {
            margin-top: 8px;
            display: flex;
            flex-direction: column;
            gap: 4px;
            font-size: 0.75em;
        }

        .inline-score-item {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 3px 8px;
            border-radius: 4px;
            background: rgba(255, 255, 255, 0.8);
            border-left: 3px solid #bdc3c7;
            transition: all 0.2s ease;
        }

        .inline-score-item:hover {
            background: rgba(255, 255, 255, 1);
            transform: translateX(2px);
        }

        .inline-score-item.positive {
            border-left-color: #27ae60;
            background: rgba(46, 204, 113, 0.1);
        }

        .inline-score-item.neutral {
            border-left-color: #f39c12;
            background: rgba(243, 156, 18, 0.1);
        }

        .inline-score-item.negative {
            border-left-color: #e74c3c;
            background: rgba(231, 76, 60, 0.1);
        }

        .inline-score-item .score-emoji {
            margin-right: 6px;
            font-size: 1.1em;
        }

        .inline-score-item .score-label {
            flex: 1;
            font-weight: 600;
            color: #2c3e50;
        }

        .inline-score-item .score-value {
            font-weight: 700;
            color: #2c3e50;
            margin-left: 8px;
        }

        /* Enhanced Sentiment Summary */
        .sentiment-summary {
            background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
            padding: 20px;
            border-radius: 15px;
            margin-bottom: 25px;
            border-left: 5px solid #17a2b8;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            position: relative;
            overflow: hidden;
        }

        .sentiment-summary::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 3px;
            background: linear-gradient(90deg, #17a2b8, #20c997, #17a2b8);
        }

        .sentiment-summary .hint-text {
            color: #6c757d;
            font-size: 0.9em;
            margin-top: 8px;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .hint-text::before {
            content: '💡';
            font-size: 1.1em;
        }

        /* Enhanced Table Row Interactions */
        .sentiment-row {
            transition: all 0.3s ease;
            position: relative;
        }

        .sentiment-row:hover {
            background: linear-gradient(135deg, #f8f9fa 0%, #fff 100%);
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            transform: translateY(-1px);
        }

        .no-analysis-reason {
            text-align: center;
            color: #6c757d;
            font-style: italic;
            padding: 20px;
            background: linear-gradient(135deg, #f8f9fa 0%, #fff 100%);
            border-radius: 8px;
            border: 2px dashed #dee2e6;
        }

        /* Section Description Styling */
        .section-description {
            color: #6c757d;
            font-size: 0.9em;
            margin-top: 8px;
            padding: 8px 12px;
            background: rgba(23, 162, 184, 0.1);
            border-radius: 6px;
            border-left: 3px solid #17a2b8;
        }

        /* Category Reference Styling (integrated into daily records section) */
        .category-reference {
            background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
            padding: 20px;
            border-radius: 12px;
            margin-bottom: 20px;
            border-left: 4px solid #6c757d;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }

        .category-reference h4 {
            display: flex;
            align-items: center;
            gap: 8px;
            margin-bottom: 15px;
            color: #2c3e50;
            font-size: 1.1em;
        }

        .category-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 12px;
        }

        .category-item {
            background: white;
            padding: 12px 16px;
            border-radius: 8px;
            border-left: 3px solid #17a2b8;
            font-size: 0.9em;
            line-height: 1.4;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            transition: all 0.2s ease;
        }

        .category-item:hover {
            transform: translateY(-1px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        }

        .category-item strong {
            color: #17a2b8;
            margin-right: 8px;
        }

        /* Responsive adjustments */
        @media (max-width: 768px) {
            .inline-confidence-details {
                font-size: 0.7em;
            }
            
            .inline-score-item {
                padding: 2px 6px;
            }

            .language-badge {
                font-size: 0.75em;
                padding: 4px 8px;
                min-width: 50px;
            }

            .category-grid {
                grid-template-columns: 1fr;
                gap: 8px;
            }
            
            .category-item {
                padding: 10px 12px;
                font-size: 0.85em;
            }
        }

        /* Print Styles */
        @media print {
            body {
                background: white;
            }
            
            .container {
                box-shadow: none;
                margin: 0;
            }
            
            .header {
                background: ${colors.headerGradient} !important;
                -webkit-print-color-adjust: exact;
            }

            .expandable-details {
                display: none;
            }
        }
    `;
}

/**
 * Gets the JavaScript code for interactive features (simplified without expandable buttons)
 * @returns {string} JavaScript code as string
 */
function getReportScripts() {
    return `
        // Initialize page
        document.addEventListener('DOMContentLoaded', function() {
            console.log('📊 Report page initialized with inline sentiment details');
        });
    `;
}

module.exports = {
    getReportStyles,
    getReportScripts
};