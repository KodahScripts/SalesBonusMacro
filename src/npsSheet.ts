class NpsSheet {
    constructor(protected workbook: ExcelScript.Workbook, protected store: Store) {
        const sheet = workbook.addWorksheet("NPS");

        sheet.getRange("A1:H1").setValues([[ "Employee #", "Salesperson", "# of Surveys", "Current Score", "90 Day Score", "Score for Bonus", "CSI Outcome", "Regional Score" ]]);
        sheet.getCell(0, 8).setValue(store.regionalScore);

        store.employees.forEach((employee, index) => {
            const row = index + 2;
            sheet.getRange(`A${row}:G${row}`).setValues([[employee.id, employee.name, employee.nps.surveys, employee.nps.current, employee.nps.average, `=IF(D${row}>E${row},D${row},E${row})`, employee.nps.outcome]])
        });
    }
}