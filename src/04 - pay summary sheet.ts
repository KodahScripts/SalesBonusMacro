class PaySummarySheet {
  constructor(workbook: ExcelScript.Workbook, employeeList: Array<Employee>) {
    const sheet = workbook.addWorksheet('Pay Summary');
    sheet.getRange("A1:P1").setValues([
      ["Employee #", "Employee Name", "Total Units", "Rank", "Spiffs to Pay", "Commission 3120", "Retro Commission", "F&I Commission", "Month End Bonus 3122", "Total EOM Bonus 8328", "Draw 3121", "Total Commission", "YTD Bucket", "Deposit Gross", "Check Column - Should be Zero", "Draw to Take"]
    ]);

    employeeList.forEach((employee, index) => {
      const row = index + 2;
      sheet.getRange(`A${row}:G${row}`).setValues([[
        employee.id,
        employee.name,
        employee.getTotalUnits(),
        `=IF(C${row}<15, "", RANK.EQ(C${row}, C:C, 0))`,
        `=IFERROR(VLOOKUP(A${row}, 'SPIFFS'!A:H, 8, 0), 0)`,
        employee.getTotalCommission('grossAmount'),
        `=IFERROR('${employee.name}'!M${employee.getResultRow(5)}, 0)`
      ]]);
    });
  }
}