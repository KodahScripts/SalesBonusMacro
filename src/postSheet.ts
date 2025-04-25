class PostSheet {
    constructor(protected workbook: ExcelScript.Workbook, protected store: Store) {
        const sheet = workbook.addWorksheet("PowerPost");

        const header = ["DATE", "GL", "AMOUNT", "EMPID", "DESC"];
        const date = new Date().toISOString().slice(0, 10);
        const todayArr = date.split('-');
        const descYear = todayArr[0].split('0')[1];
        const descDate = [todayArr[1], descYear].join('');
        const today = Number([todayArr[1], todayArr[2], descYear].join(''));
        const employeeCount = store.employees.length;

        sheet.getRange('A1:E1').setValues([header]);

        let startRow = 2;
        store.employees.forEach((employee, index) => {
            const row = index + startRow;
            let remainder = employee.getRemainder();
            remainder = Number(remainder.toFixed(2));
            sheet.getRange(`A${row}:E${row}`).setValues([[
                `=TEXT(${today}, "000000")`,
                store.accounts.retro,
                remainder * -100,
                employee.id,
                `SLSBONUS${descDate}`
            ]]);
        });

        startRow += employeeCount;
        store.employees.forEach((employee, index) => {
            const row = index + startRow;
            const expense = Number(employee.expense.one.toFixed(2));
            sheet.getRange(`A${row}:E${row}`).setValues([[
                `=TEXT(${today}, "000000")`,
                store.accounts.expense1,
                expense * 100,
                employee.id,
                `EXPENSE1${descDate}`
            ]]);
        });

        startRow += employeeCount;
        store.employees.forEach((employee, index) => {
            const row = index + startRow;
            const expense = Number(employee.expense.two.toFixed(2));
            sheet.getRange(`A${row}:E${row}`).setValues([[
                `=TEXT(${today}, "000000")`,
                store.accounts.expense2,
                expense * 100,
                employee.id,
                `EXPENSE2${descDate}`
            ]]);
        });
    }
}