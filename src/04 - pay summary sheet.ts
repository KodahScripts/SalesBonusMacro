class PaySummarySheet {
  constructor(workbook: ExcelScript.Workbook, store: Store) {
    const sheet = workbook.addWorksheet('Pay Summary');
    sheet.getRange("A1:Q1").setValues([
      ["Employee #", "Employee Name", "Total Units", "Rank", "Finance Income Totals", "Spiffs to Pay", "Commission 3120", "Retro Commission", "F&I Commission", "Month End Bonus 3122", "Total EOM Bonus 8328", "Draw 3121", "Total Commission", "YTD Bucket", "Deposit Gross", "Check Column - Should be Zero", "Draw to Take"]
    ]);
    store.employees.forEach((employee, index) => {
      const row = index + 2;
      sheet.getRange(`A${row}:J${row}`).setValues([[
        employee.id,
        employee.name,
        employee.unitCount,
        `=IF(C${row}<15, "", RANK.EQ(C${row}, C:C, 0))`,
        employee.commissionTotals.fni,
        employee.spiffs,
        employee.commissionTotals.amount,
        employee.retroTotal,
        employee.fniTotal.payout,
        employee.bonus.total,
      ]]);
    });
  }
}