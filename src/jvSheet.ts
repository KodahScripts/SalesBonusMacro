class JvSheet {
    constructor(protected workbook: ExcelScript.Workbook, protected store: Store) {
        const sheet = workbook.addWorksheet("JV Posting");

        sheet.getRange("A1:M1").setValues([[ "Employee #", "Employee Name", "Draw", "Commission", "Retro Commission", "F&I Commission", "Bonus", "Spiffs", "Total Commission / Bonus", "Total Due / Owed", "YTD Bucket", "Expense 1", "Expense 2" ]]);

        store.employees.forEach((employee, index) => {
            const row = index + 2;
            const owed = employee.getOwed();
            const newPercent = (store.units.new / store.units.total) * owed;
            const usedPercent = (store.units.used / store.units.total) * owed;

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
                newPercent,
                usedPercent
             ]]);
        });

        const lastRow = store.employees.length + 5;
        sheet.getRange(`B${lastRow}:D${lastRow+2}`).setValues([
            ["UNITS", store.units.total, ''],
            ["NEW", store.units.new, `${(store.units.new / store.units.total) * 100}%`],
            ["USED", store.units.used, `${(store.units.used / store.units.total) * 100}%`]
        ]);
    }
}