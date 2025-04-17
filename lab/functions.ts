const STORE_NAME = "BMW of South Miami";
const STORE_ABBR = "BOSM";

function main(workbook: ExcelScript.Workbook) {
    const initialSheets: Array<ExcelScript.Worksheet> = workbook.getWorksheets();

    initialSheets.forEach(sheet => {
        const reportData = sheet.getUsedRange().getValues();

        switch (sheet.getName()) {
            case '0432': 
                                
                break;
            case '90': data_90day = sheet.getUsedRange().getValues();
                break;
            case '3213': data_3213 = sheet.getUsedRange().getValues();
                break;
            case 'SPIFFS': data_spiffs = sheet.getUsedRange().getValues();
                break;
            case 'NPS Sheet': data_nps = sheet.getUsedRange().getValues();
                break;
            case 'Look Up Table': data_lut = sheet.getUsedRange().getValues();
                break;
            default: sheet.delete()
                break;
        }
    })
}


