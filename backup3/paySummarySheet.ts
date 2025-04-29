// class PaySummarySheet {
//     private sheet: ExcelScript.Worksheet;
//     constructor(protected workbook: ExcelScript.Workbook, protected store: Store) {
//         this.sheet = workbook.addWorksheet("Pay Summary");
//         this.sheet.getRange("A1:T1").setValues([["Employee #", "Salesperson", "Total Units", "Gross for Rank", "Bonus Rank", "F&I Totals", "Spiffs to Pay", "Commission 3120", "Retro Commission", "F&I Commission", "Month End Bonus 3122", "Total EOM Bonus 8328", "Draw 3121", "Total Commissions", "YTD Bucket", "Deposit Gross", "Check - Should be 0", "Draw to Take", "Expense 1", "Expense 2"]]);

//         const startRow = 2

//         store.employees.forEach((employee, index) => {
//             const row = index + startRow;
//             const rankFormula = `=IF(C${row}<15,"",RANK.EQ(C${row},C:C,0))`;
//             const depositGross = employee.getDepositGross();
//             const balanceComm = `=H${row} + ${employee.commissionBalance}`;

//             this.sheet.getRange(`A${row}:T${row}`).setValues([[
//                 employee.id,
//                 employee.name,
//                 employee.units.total,
//                 employee.commission.gross.toFixed(2),
//                 rankFormula,
//                 employee.commission.fni.toFixed(2),
//                 employee.spiff.toFixed(2),
//                 employee.commission.amount.toFixed(2),
//                 employee.retro.total.toFixed(2),
//                 employee.fni.payout.toFixed(2),
//                 employee.bonus.total.toFixed(2),
//                 employee.bonus.eom.toFixed(2),
//                 employee.priorDraw.toFixed(2),
//                 employee.totalCommission.toFixed(2),
//                 employee.ytdBucket.toFixed(2),
//                 depositGross.toFixed(2),
//                 balanceComm,
//                 employee.drawAmount.toFixed(2),
//                 employee.expense.one.toFixed(2),
//                 employee.expense.two.toFixed(2)
//             ]]);
//         });

//         const reportEnd = store.employees.length + 1;
//         let totalRow = reportEnd + 1;
//         this.sheet.getRange(`C${totalRow}:T${totalRow}`).setValues([[
//             `=ROUND(SUM(C${startRow}:C${reportEnd}), 2)`,
//             `=ROUND(SUM(D${startRow}:D${reportEnd}), 2)`,
//             '',
//             `=ROUND(SUM(F${startRow}:F${reportEnd}), 2)`,
//             `=ROUND(SUM(G${startRow}:G${reportEnd}), 2)`,
//             `=ROUND(SUM(H${startRow}:H${reportEnd}), 2)`,
//             `=ROUND(SUM(I${startRow}:I${reportEnd}), 2)`,
//             `=ROUND(SUM(J${startRow}:J${reportEnd}), 2)`,
//             `=ROUND(SUM(K${startRow}:K${reportEnd}), 2)`,
//             `=ROUND(SUM(L${startRow}:L${reportEnd}), 2)`,
//             `=ROUND(SUM(M${startRow}:M${reportEnd}), 2)`,
//             `=ROUND(SUM(N${startRow}:N${reportEnd}), 2)`,
//             `=ROUND(SUM(O${startRow}:O${reportEnd}), 2)`,
//             `=ROUND(SUM(P${startRow}:P${reportEnd}), 2)`,
//             `=ROUND(SUM(Q${startRow}:Q${reportEnd}), 2)`,
//             `=ROUND(SUM(R${startRow}:R${reportEnd}), 2)`,
//             `=ROUND(SUM(R${startRow}:S${reportEnd}), 2)`,
//             `=ROUND(SUM(R${startRow}:T${reportEnd}), 2)`,
//         ]]);

//         const belowAllStartRow = totalRow + 2;
//         this.sheet.getRange(`B${belowAllStartRow}:C${belowAllStartRow + 1}`).setValues([
//             ["NEW", `${store.units.new} (${(store.unitPercent.new * 100).toFixed(1)}%)`],
//             ["USED", `${store.units.used} (${(store.unitPercent.used * 100).toFixed(1)}%)`]
//         ]);
//         this.format();
//     }
//     format() {
//         const employeeCount = this.store.employees.length + 1;
//         const data = this.sheet.getRange(`A1:T${employeeCount}`);
//         const dataTable = this.sheet.addTable(data, true);

//         dataTable.setPredefinedTableStyle("TableStyleLight2");
//     }
// }