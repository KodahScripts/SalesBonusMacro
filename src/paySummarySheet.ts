class PaySummarySheet {
    constructor(protected workbook: ExcelScript.Workbook, protected store: Store) {
        const sheet = workbook.addWorksheet("Pay Summary");

        sheet.getRange("A1:R1").setValues([[ "Employee #", "Salesperson", "Total Units", "Gross for Rank", "Bonus Rank", "F&I Totals", "Spiffs to Pay", "Commission 3120", "Retro Commission", "F&I Commission", "Month End Bonus 3122", "Total EOM Bonus 8328", "Draw 3121", "Total Commissions", "YTD Bucket", "Deposit Gross", "Check - Should be 0", "Draw to Take" ]]);

        store.employees.forEach((employee, index) => {
            const row = index + 2;
            const rankFormula = `=IF(C${row}<15,"",RANK.EQ(C${row},C:C,0))`;
            const depositGross = employee.getDepositGross();
            const balanceComm = `=H${row} + ${employee.commissionBalance}`;

            sheet.getRange(`A${row}:R${row}`).setValues([[
                employee.id, 
                employee.name, 
                employee.units.total, 
                employee.commission.gross, 
                rankFormula, 
                employee.commission.fni, 
                employee.spiff, 
                employee.commission.amount, 
                employee.retro.total, 
                employee.fni.payout, 
                employee.bonus.total, 
                employee.bonus.eom, 
                employee.priorDraw,
                employee.totalCommission,
                employee.ytdBucket,
                depositGross,
                balanceComm,
                employee.drawAmount
            ]]);
        });
    }
}