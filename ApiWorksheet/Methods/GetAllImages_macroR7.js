/**
 * R7 Office JavaScript макрос - ApiWorksheet.GetAllImages
 * 
 *  Демонстрация использования метода GetAllImages класса ApiWorksheet
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
        // This example shows how to get all images from the sheet.
        
        // How to get all images.
        
        // Get all images as an array.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.AddImage("https://api.R7 Office.com/content/img/docbuilder/examples/coordinate_aspects.png", 60 * 36000, 35 * 36000, 0, 2 * 36000, 0, 3 * 36000);
        let images = worksheet.GetAllImages();
        let classType = images[0].GetClassType();
        worksheet.GetRange("A10").SetValue("Class Type = " + classType);
        
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
