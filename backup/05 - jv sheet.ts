class JvSheet {
  constructor(workbook: ExcelScript.Workbook, employeeList: Array<Employee>) {
    const sheet = workbook.addWorksheet('JV Posting');
    sheet.getRange("A1:M1").setValues([
      ["Employee #", "Employee Name", "Draw", "Commission", "Retro Commission", "F&I Commission", "Bonus", "Spiffs", "Total Comm/Bonus", "Total Due/Owed", "YTD Bucket", "Expense 1", "Expense 2"]
    ]);

    employeeList.forEach((employee, index) => {
      const row = index + 2;
      sheet.getRange(`A${row}:B${row}`).setValues([[employee.id, employee.name]])
    });
  }
}