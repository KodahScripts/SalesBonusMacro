class JvSheet {
    constructor(protected workbook: ExcelScript.Workbook, protected store: Store) {
        const sheet = workbook.addWorksheet("JV Posting");

        sheet.getRange("A1:M1").setValues([["Employee #", "Employee Name", "Draw", "Commission", "Retro Commission", "F&I Commission", "Bonus", "Spiffs", "Total Commission / Bonus", "Total Due / Owed", "YTD Bucket", "Expense 1", "Expense 2"]]);

        const startRow = 2

        store.employees.forEach((employee, index) => {
            const row = index + startRow;
            const owed = employee.getOwed();

            sheet.getRange(`A${row}:M${row}`).setValues([[
                employee.id,
                employee.name,
                employee.priorDraw,
                employee.commission.amount,
                employee.retro.total,
                employee.fni.payout,
                employee.bonus.total,
                employee.spiff,
                employee.totalCommission,
                owed,
                employee.ytdBucket,
                employee.expense.one,
                employee.expense.two
            ]]);
        });

        const reportEnd = store.employees.length + 1;
        let totalRow = reportEnd + 1;
        sheet.getRange(`C${totalRow}:M${totalRow}`).setValues([[
            `=SUM(C${startRow}:C${reportEnd})`,
            `=SUM(D${startRow}:D${reportEnd})`,
            `=SUM(E${startRow}:E${reportEnd})`,
            `=SUM(F${startRow}:F${reportEnd})`,
            `=SUM(G${startRow}:G${reportEnd})`,
            `=SUM(H${startRow}:H${reportEnd})`,
            `=SUM(I${startRow}:I${reportEnd})`,
            `=SUM(J${startRow}:J${reportEnd})`,
            `=SUM(K${startRow}:K${reportEnd})`,
            `=SUM(L${startRow}:L${reportEnd})`,
            `=SUM(M${startRow}:M${reportEnd})`
        ]]);

        totalRow++;
        sheet.getRange(`C${totalRow}:M${totalRow}`).setValues([[
            store.totalDraw,
            store.commission.amount,
            store.retro.total,
            store.fni.payout,
            store.bonus.total,
            store.totalSpiffs,
            store.totalCommission,
            store.totalOwed,
            store.ytdBucket,
            store.expense.one,
            store.expense.two
        ]]);

        const calcRow = totalRow + 1;
        sheet.getRange(`C${calcRow}:M${calcRow}`).setValues([[
            `=C${calcRow - 2}-C${calcRow - 1}`,
            `=D${calcRow - 2}-D${calcRow - 1}`,
            `=E${calcRow - 2}-E${calcRow - 1}`,
            `=F${calcRow - 2}-F${calcRow - 1}`,
            `=G${calcRow - 2}-G${calcRow - 1}`,
            `=H${calcRow - 2}-H${calcRow - 1}`,
            `=I${calcRow - 2}-I${calcRow - 1}`,
            `=J${calcRow - 2}-J${calcRow - 1}`,
            `=K${calcRow - 2}-K${calcRow - 1}`,
            `=L${calcRow - 2}-L${calcRow - 1}`,
            `=M${calcRow - 2}-M${calcRow - 1}`
        ]]);

        const belowAllStartRow = calcRow + 3;
        sheet.getRange(`B${belowAllStartRow}:D${belowAllStartRow + 2}`).setValues([
            ["UNITS", store.units.total, ''],
            ["NEW", store.units.new, store.unitPercent.new * 100],
            ["USED", store.units.used, store.unitPercent.used * 100]
        ]);
    }
}