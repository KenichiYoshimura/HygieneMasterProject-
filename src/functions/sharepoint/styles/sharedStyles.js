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
            vertical-align: middle;
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

        .comment-cell {
            text-align: left;
            max-width: 200px;
            font-size: 0.9em;
            color: #495057;
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

        /* Confidence Tooltip Styles */
        .confidence-tooltip {
            position: relative;
        }

        .confidence-details {
            position: absolute;
            bottom: 100%;
            left: 50%;
            transform: translateX(-50%);
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
            color: white;
            padding: 12px 16px;
            border-radius: 10px;
            font-size: 0.85em;
            line-height: 1.6;
            box-shadow: 0 8px 25px rgba(0,0,0,0.3);
            border: 1px solid #34495e;
            opacity: 0;
            visibility: hidden;
            transform: translateX(-50%) translateY(-10px);
            transition: all 0.3s cubic-bezier(0.68, -0.55, 0.265, 1.55);
            z-index: 1000;
            min-width: 180px;
            white-space: nowrap;
        }

        .confidence-details::before {
            content: '';
            position: absolute;
            top: 100%;
            left: 50%;
            transform: translateX(-50%);
            border: 8px solid transparent;
            border-top-color: #2c3e50;
        }

        .confidence-score-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin: 4px 0;
            padding: 2px 0;
        }

        .score-emoji {
            margin-right: 8px;
            font-size: 1.1em;
        }

        .score-label {
            flex: 1;
            text-align: left;
        }

        .score-value {
            font-weight: 700;
            margin-left: 8px;
        }

        .confidence-score-item.positive .score-value {
            color: #2ecc71;
        }

        .confidence-score-item.neutral .score-value {
            color: #f39c12;
        }

        .confidence-score-item.negative .score-value {
            color: #e74c3c;
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

        .sentiment-row:hover .confidence-details {
            opacity: 1;
            visibility: visible;
            transform: translateX(-50%) translateY(-5px);
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

        /* Enhanced Expandable Details Styles with debugging */
        .expandable-details {
            margin-top: 8px;
            width: 100%;
            position: relative;
        }

        .details-toggle {
            background: linear-gradient(135deg, #3498db 0%, #2980b9 100%) !important;
            color: white !important;
            border: none !important;
            padding: 8px 12px !important;
            border-radius: 6px !important;
            font-size: 0.8em !important;
            font-weight: 600 !important;
            cursor: pointer !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            gap: 6px !important;
            transition: all 0.2s ease !important;
            width: 100% !important;
            box-shadow: 0 2px 4px rgba(52, 152, 219, 0.3) !important;
            min-height: 32px !important;
            outline: none !important;
            position: relative !important;
            z-index: 10 !important;
        }

        .details-toggle:hover {
            background: linear-gradient(135deg, #2980b9 0%, #1f5f8b 100%) !important;
            transform: translateY(-1px) !important;
            box-shadow: 0 4px 8px rgba(52, 152, 219, 0.4) !important;
        }

        .details-toggle:active {
            transform: translateY(0) !important;
            background: linear-gradient(135deg, #1f5f8b 0%, #1a4971 100%) !important;
        }

        .details-toggle:focus {
            outline: 2px solid #3498db !important;
            outline-offset: 2px !important;
        }

        .details-toggle[aria-expanded="true"] {
            background: linear-gradient(135deg, #27ae60 0%, #229954 100%) !important;
            box-shadow: 0 2px 4px rgba(39, 174, 96, 0.3) !important;
        }

        .details-toggle[aria-expanded="true"]:hover {
            background: linear-gradient(135deg, #229954 0%, #1e8449 100%) !important;
        }

        .toggle-icon {
            transition: transform 0.3s ease !important;
            font-size: 0.7em !important;
            display: inline-block !important;
        }

        .details-toggle[aria-expanded="true"] .toggle-icon {
            transform: rotate(180deg) !important;
        }

        .details-content {
            background: linear-gradient(135deg, #f8f9fa 0%, #ffffff 100%) !important;
            border: 2px solid #e9ecef !important;
            border-radius: 8px !important;
            padding: 15px !important;
            margin-top: 8px !important;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1) !important;
            animation: slideDown 0.3s ease !important;
            position: relative !important;
            z-index: 5 !important;
        }

        @keyframes slideDown {
            from {
                opacity: 0;
                transform: translateY(-10px);
                max-height: 0;
            }
            to {
                opacity: 1;
                transform: translateY(0);
                max-height: 200px;
            }
        }

        .details-header {
            color: #2c3e50 !important;
            font-size: 0.9em !important;
            margin-bottom: 12px !important;
            padding-bottom: 8px !important;
            border-bottom: 2px solid #bdc3c7 !important;
            text-align: center !important;
            font-weight: bold !important;
        }

        /* Debug styles */
        .details-toggle::before {
            content: '';
            position: absolute;
            top: -2px;
            left: -2px;
            right: -2px;
            bottom: -2px;
            background: transparent;
            border: 1px dashed red;
            opacity: 0;
            pointer-events: none;
        }

        /* Show debug border on hover (remove in production) */
        .details-toggle:hover::before {
            opacity: 0.3;
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
        }

        /* Additional styles for sentiment analysis and error handling */
        .no-translation {
            background: #e8f5e8;
            color: #2e7d32;
            padding: 4px 8px;
            border-radius: 12px;
            font-size: 0.8em;
            font-weight: 500;
        }

        .sentiment-summary {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
            border-left: 4px solid #17a2b8;
        }

        .no-analysis-reason {
            text-align: center;
            color: #666;
            font-style: italic;
            padding: 15px;
        }

        .analysis-error {
            text-align: center;
            color: #666;
            font-style: italic;
        }

        .error-badge {
            background: #f8d7da;
            color: #721c24;
            padding: 4px 8px;
            border-radius: 12px;
            font-size: 0.8em;
            font-weight: 600;
        }

        .sentiment-row.error {
            background: #fff8f0;
        }

        .sentiment-row.no-analysis {
            background: #f8f9fa;
        }
    `;
}

/**
 * Gets the JavaScript code for interactive features
 * @returns {string} JavaScript code as string
 */
function getReportScripts() {
    return `
        // Toggle expandable details sections
        function toggleDetails(recordId) {
            console.log('Toggling details for:', recordId);
            
            const detailsElement = document.getElementById('details-' + recordId);
            const toggleButton = document.querySelector('[onclick*="' + recordId + '"]');
            const toggleText = toggleButton ? toggleButton.querySelector('.toggle-text') : null;
            
            if (!detailsElement) {
                console.error('Details element not found for:', recordId);
                return;
            }
            
            console.log('Current display:', detailsElement.style.display);
            
            if (detailsElement.style.display === 'none' || detailsElement.style.display === '') {
                detailsElement.style.display = 'block';
                if (toggleButton) {
                    toggleButton.setAttribute('aria-expanded', 'true');
                }
                if (toggleText) {
                    toggleText.textContent = '閉じる';
                }
                console.log('Opened details for:', recordId);
            } else {
                detailsElement.style.display = 'none';
                if (toggleButton) {
                    toggleButton.setAttribute('aria-expanded', 'false');
                }
                if (toggleText) {
                    toggleText.textContent = '詳細';
                }
                console.log('Closed details for:', recordId);
            }
        }

        // Alternative event listener approach for better reliability
        document.addEventListener('DOMContentLoaded', function() {
            console.log('📊 Report page initialized with expandable sentiment details');
            
            // Add click listeners to all detail toggle buttons
            const toggleButtons = document.querySelectorAll('.details-toggle');
            console.log('Found', toggleButtons.length, 'toggle buttons');
            
            toggleButtons.forEach(function(button) {
                button.addEventListener('click', function(event) {
                    event.preventDefault();
                    event.stopPropagation();
                    
                    const onclick = button.getAttribute('onclick');
                    if (onclick) {
                        const recordIdMatch = onclick.match(/toggleDetails\\('([^']+)'\\)/);
                        if (recordIdMatch) {
                            const recordId = recordIdMatch[1];
                            toggleDetails(recordId);
                        }
                    }
                });
            });
        });

        // Close all expanded details when clicking outside
        document.addEventListener('click', function(event) {
            if (!event.target.closest('.expandable-details')) {
                const allDetails = document.querySelectorAll('.details-content');
                const allToggleButtons = document.querySelectorAll('.details-toggle');
                
                allDetails.forEach(function(detail) {
                    if (detail.style.display === 'block') {
                        detail.style.display = 'none';
                    }
                });
                
                allToggleButtons.forEach(function(button) {
                    button.setAttribute('aria-expanded', 'false');
                    const toggleText = button.querySelector('.toggle-text');
                    if (toggleText) {
                        toggleText.textContent = '詳細';
                    }
                });
            }
        });

        // Keyboard accessibility
        document.addEventListener('keydown', function(event) {
            if (event.key === 'Escape') {
                const allDetails = document.querySelectorAll('.details-content[style*="block"]');
                const allToggleButtons = document.querySelectorAll('.details-toggle[aria-expanded="true"]');
                
                allDetails.forEach(function(detail) {
                    detail.style.display = 'none';
                });
                
                allToggleButtons.forEach(function(button) {
                    button.setAttribute('aria-expanded', 'false');
                    const toggleText = button.querySelector('.toggle-text');
                    if (toggleText) {
                        toggleText.textContent = '詳細';
                    }
                });
            }
        });
    `;
}

// Make sure this file exists and is properly structured
module.exports = {
    getReportStyles,
    getReportScripts};