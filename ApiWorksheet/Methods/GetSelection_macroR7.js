/**
 * R7 Office JavaScript макрос - ApiWorksheet.GetSelection
 * 
 *  Демонстрация использования метода GetSelection класса ApiWorksheet
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
        // This example shows how to get an object that represents the selected range.
        
        // How to get selected range.
        
        // Get selection from the worksheet and set its value.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetSelection().SetValue("selected");
        
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
