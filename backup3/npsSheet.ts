// class NpsSheet {
//     private sheet: ExcelScript.Worksheet;
//     constructor(protected workbook: ExcelScript.Workbook, protected store: Store) {
//         this.sheet = workbook.addWorksheet("NPS");
//         this.store = store;

//         this.sheet.getRange("A1:H1").setValues([["Employee #", "Salesperson", "# of Surveys", "Current Score", "90 Day Score", "Score for Bonus", "CSI Outcome", "Regional Score"]]);
//         this.sheet.getCell(0, 8).setValue(store.regionalScore);

//         store.employees.forEach((employee, index) => {
//             const row = index + 2;
//             this.sheet.getRange(`A${row}:G${row}`).setValues([[employee.id, employee.name, employee.nps.surveys, employee.nps.current / 100, employee.nps.average / 100, `=IF(D${row}>E${row},D${row},E${row})`, employee.nps.outcome]])
//         });
//         this.format();
//     }
//     format() {
//         const employeeCount = this.store.employees.length + 1;
//         const data = this.sheet.getRange(`A1:G${employeeCount}`);
//         const dataTable = this.sheet.addTable(data, true);
//         dataTable.setPredefinedTableStyle("TableStyleLight2");
//         const regionalRange = this.sheet.getRange("H1:I1").getFormat();

//         this.sheet.getRange(`D2:F${employeeCount}`).setNumberFormat("0.0%");

//         regionalRange.setColumnWidth(80);
//         regionalRange.setRowHeight(50);
//         regionalRange.getFont().setBold(true);
//         regionalRange.getFill().setColor("lightblue");

//         this.sheet.getRange("A:A").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
//         this.sheet.getRange("A:A").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
//         this.sheet.getRange("C:I").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
//         this.sheet.getRange("C:I").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
//         data.getFormat().autofitColumns();
//     }
// }