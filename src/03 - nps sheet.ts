class NpsSheet {
  constructor(workbook: ExcelScript.Workbook, employeeList: Array<Employee>) {
    const sheet = workbook.addWorksheet('NPS');
    sheet.getRange("A1:H1").setValues([
      ["Regional Score", "Employee #", "Employee Name", "# of Surveys", "Monthly Score", "90 Day Score", "NPS Score for Bonus", "CSI Outcome"]
    ]);
    employeeList.forEach((employee, index) => {
      const row = index + 2;
      sheet.getRange(`B${row}:F${row}`).setValues([[employee.id, employee.name, employee.nps.surveys, employee.nps.curr_percent, employee.nps.avg_percent]]);
      sheet.getRange(`G${row}`).setValue(`=IF(E${row}>F${row},E${row},F${row})`)
      sheet.getRange(`H${row}`).setValue(`=IF(G${row}>A2+3%,"3P",IF(G${row}=A2,"A",IF(G${row}<A2,"B")))`);
    });
    sheet.getRange("1:1").getFormat().getFill().setColor("lightgrey");
    sheet.getRange("1:1").getFormat().getFont().setBold(true);
    sheet.getRange("1:1").getFormat().setColumnWidth(120);
    sheet.getRange("1:1").getFormat().setRowHeight(50);
    sheet.getRange("A2:A3").merge();
    sheet.getCell(1, 0).getFormat().getFill().setColor("yellow");
    sheet.getCell(1, 0).setNumberFormat("0.0%");
    sheet.getRange("E2:G100").setNumberFormat('0%');
    sheet.getRange("A1:H100").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    sheet.getRange("A1:H100").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  }
}