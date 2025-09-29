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

        .confidence-container {
            position: relative;
            background: #e9ecef;
            border-radius: 10px;
            height: 20px;
            min-width: 60px;
        }

        .confidence-fill {
            height: 100%;
            border-radius: 10px;
            transition: width 0.3s ease;
        }

        .confidence-text {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            font-size: 0.75em;
            font-weight: 600;
            color: #2c3e50;
        }

        .language-tag {
            background: #e3f2fd;
            color: #1976d2;
            padding: 4px 8px;
            border-radius: 12px;
            font-size: 0.8em;
            font-weight: 500;
        }

        .translation-text, .comment-text {
            text-align: left;
            max-width: 250px;
            font-size: 0.9em;
            line-height: 1.4;
        }

        /* Footer */
        .footer {
            text-align: center;
            padding: 20px;
            background: #f8f9fa;
            color: #6c757d;
            border-radius: 0 0 12px 12px;
            margin: 30px -20px -20px -20px;
            font-size: 0.9em;
        }

        .footer .timestamp {
            font-weight: 600;
            color: #495057;
        }

        /* Responsive Design */
        @media (max-width: 768px) {
            .container {
                margin: 10px;
                padding: 15px;
            }
            
            .summary-cards {
                grid-template-columns: 1fr;
            }
            
            .header h1 {
                font-size: 1.8em;
            }
            
            table {
                font-size: 0.85em;
            }
            
            .section-content {
                padding: 20px;
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
        // Add interactivity for collapsible sections
        document.querySelectorAll('.section-header').forEach(header => {
            header.addEventListener('click', () => {
                const content = header.nextElementSibling;
                const icon = header.querySelector('.toggle-icon');
                
                if (content.style.display === 'none') {
                    content.style.display = 'block';
                    if (icon) icon.style.transform = 'rotate(0deg)';
                } else {
                    content.style.display = 'none';
                    if (icon) icon.style.transform = 'rotate(-90deg)';
                }
            });
        });
    `;
}

// Make sure this file exists and is properly structured
module.exports = {
    getReportStyles,
    getReportScripts
};