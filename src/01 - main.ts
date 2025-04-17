function main(workbook:ExcelScript.Workbook) {
    const initialSheets: Array<ExcelScript.Worksheet> = workbook.getWorksheets();

    let data_0432  : Array<string | number | boolean>[] = []
    let data_90day : Array<string | number | boolean>[] = []
    let data_3213  : Array<string | number | boolean>[] = []
    let data_spiffs: Array<string | number | boolean>[] = []
    let data_nps   : Array<string | number | boolean>[] = []
    let data_lut   : Array<string | number | boolean>[] = []

    initialSheets.forEach(sheet => {
    switch (sheet.getName()) {
        case '0432': data_0432 = sheet.getUsedRange().getValues();
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

    const store = new Store(data_lut);
    store.createEmployees(data_0432);

    new NpsSheet(workbook, store.employees);

    new PaySummarySheet(workbook, store.employees);
    new JvSheet(workbook, store.employees);

    store.employees.forEach(employee => {
        new SalesSheet(workbook, employee)
    });
}