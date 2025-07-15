/**
 * R7 Office JavaScript макрос - ApiWorksheet.SetHyperlink
 * 
 *  Демонстрация использования метода SetHyperlink класса ApiWorksheet
 * https://r7-consult.ru/
 */

(function() {
    'use strict';
    
    try {
        // Initialize R7 Office API
        const api = Api;
        if (!api) {
            throw new Error('R7 Office API not available');
        }
        
        // Original code enhanced with error handling:
        // This example adds a hyperlink to the specified range.
        
        // How to add hyperlinks to the range.
        
        // Add a hyperlink to the cell.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetHyperlink("A1", "https://api.R7 Office.com/docbuilder/basic", "Api R7 Office", "R7 Office for developers");
        
        // Success notification
        console.log('Macro executed successfully');
        
    } catch (error) {
        console.error('Macro execution failed:', error.message);
        // Optional: Show error to user
        if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
            const sheet = Api.GetActiveSheet();
            if (sheet) {
                sheet.GetRange('A1').SetValue('Error: ' + error.message);
            }
        }
    }
})();
