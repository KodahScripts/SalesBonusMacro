class PaySheet {
    constructor(protected workbook: ExcelScript.Workbook, protected store: Store) {
        const sheet = workbook.addWorksheet("Pay Check");

        sheet.getRange("A1:H1").setValues([["Acct", "Amount", "Control", "Acct Number", "SLSBN", "Score for Bonus", "CSI Outcome", "Regional Score"]]);
        sheet.getCell(0, 8).setValue(store.regionalScore);

        store.employees.forEach((employee, index) => {
            const row = index + 2;
            sheet.getRange(`A${row}:G${row}`).setValues([[]])
        });
    }
}