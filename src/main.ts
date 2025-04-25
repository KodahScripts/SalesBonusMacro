function main(workbook: ExcelScript.Workbook) {
    const initialSheets = workbook.getWorksheets();
    const store = new Store("BMW of South Miami", "BOSM");

    initialSheets.forEach(sheet => {
        const data = sheet.getUsedRange().getValues();
        let header = data.shift();
        const sheetName = sheet.getName();
        switch (sheetName) {
            case 'INPUT':
                const storeName_input = header.indexOf("Store Name");
                const storeAbbr_input = header.indexOf("Store Abbr");
                const date_input = header.indexOf("Date");
                const regionalScore_input = header.indexOf("Regional Score");
                const retroAcct_input = header.indexOf("Retro Acct");
                const expenseAcct1_input = header.indexOf("Expense 1 Acct");
                const expenseAcct2_input = header.indexOf("Expense 2 Acct");
                const saleTaxAcct_input = header.indexOf("Sales Tax Acct");
                const saleBonusTaxAcct1_input = header.indexOf("Sales Bonus Tax 1");
                const saleBonusTaxAcct2_input = header.indexOf("Sales Bonus Tax 2");
                data.forEach(row => {
                    store.name = String(row[storeName_input]);
                    store.abbr = String(row[storeAbbr_input]);
                    store.date = String(row[date_input]);
                    store.regionalScore = Number(row[regionalScore_input]);
                    store.accounts.retro = String(row[retroAcct_input]);
                    store.accounts.expense1 = String(row[expenseAcct1_input]);
                    store.accounts.expense2 = String(row[expenseAcct2_input]);
                    store.accounts.salesTax = String(row[saleTaxAcct_input]);
                    store.accounts.salesBonusTax1 = String(row[saleBonusTaxAcct1_input]);
                    store.accounts.salesBonusTax2 = String(row[saleBonusTaxAcct2_input]);
                });
                break;
            case '0432':
                const empID_0432 = header.indexOf("Salesperson#");
                const empName_0432 = header.indexOf("Salesperson Name");
                const dealID_0432 = header.indexOf("Reference#");
                const dealDate_0432 = header.indexOf("Date");
                const custID_0432 = header.indexOf("Customer#");
                const custName_0432 = header.indexOf("Customer Name");
                const vehID_0432 = header.indexOf("Stock#");
                const vehDesc_0432 = header.indexOf("Description");
                const saleType_0432 = header.indexOf("Sale Type");
                const commFni_0432 = header.indexOf("COMMBL F&I");
                const commGross_0432 = header.indexOf("COMMBL FRONT");
                const units_0432 = header.indexOf("Units");
                const commAmount_0432 = header.indexOf("Commission Amount");
                data.forEach(row => {
                    const empID = Number(row[empID_0432]);
                    const unitCount = Number(row[units_0432]);
                    if (!store.employeeExists(empID)) store.employees.push(new Employee(empID, String(row[empName_0432])));
                    const employee = store.employees.find(emp => emp.id === empID);
                    if (unitCount > 0) {
                        const customer: Person = { id: Number(row[custID_0432]), name: String(row[custName_0432]) };
                        const vehicle = new Vehicle(String(row[vehID_0432]), String(row[vehDesc_0432]), String(row[saleType_0432]));
                        const commission: Commission = { fni: Number(row[commFni_0432]), gross: Number(row[commGross_0432]), amount: Number(row[commAmount_0432]) };
                        const deal = new Deal(String(row[dealID_0432]), Number(row[dealDate_0432]), customer, vehicle, unitCount, commission);
                        employee?.deals.push(deal);
                    }
                });
                break;
            case '90':
                const empID_90 = header.indexOf("Salesperson#");
                const units_90 = header.indexOf("units");
                data.forEach(row => {
                    const empID = Number(row[empID_90]);
                    const unitCount = Number(row[units_90]);
                    const employee = store.employees.find(emp => emp.id === empID);
                    employee?.setAverageUnits(unitCount);
                });
                break;
            case '3213':
                const empID_3213 = header.indexOf("Control#");
                const comm_3213 = header.indexOf("8321C");
                const amount_3213 = header.indexOf("8321D");
                data.forEach(row => {
                    const empID = Number(row[empID_3213]);
                    const commissionBalance = Number(row[comm_3213]);
                    const amount = Number(row[amount_3213]);
                    const employee = store.employees.find(emp => emp.id === empID);
                    employee?.priorDraw = amount;
                    employee?.commissionBalance = commissionBalance;
                });
                break;
            case 'SPIFFS':
                const empID_spiff = header.indexOf("Employee #");
                const amount_spiff = header.indexOf("Total");
                data.forEach(row => {
                    const empID = Number(row[empID_spiff]);
                    const amount = Number(row[amount_spiff]);
                    const employee = store.employees.find(emp => emp.id === empID);
                    employee?.spiff = amount;
                });
                break;
            case 'NPS Sheet':
                header = data[1];
                const empID_nps = header.indexOf("Employee #");
                const surveyCount_nps = header.indexOf("PROMOTER");
                const current_nps = 8;
                const average_nps = 23;
                data.forEach(row => {
                    const empID = Number(row[empID_nps]);
                    const surveys = Number(row[surveyCount_nps]);
                    const current = Number(row[current_nps]) * 100;
                    const average = Number(row[average_nps]) * 100;
                    const outcome = calculateCsiOutcome(current, average, store.regionalScore);
                    const employee = store.employees.find(emp => emp.id === empID);
                    employee?.nps = { surveys, current, average, outcome };
                });
                break;
            case 'Look Up Table':
                break;
            default: sheet.delete();
                break;
        }
    });
    store.calculateAll()

    const allSheets = [
        new NpsSheet(workbook, store),
        new PaySummarySheet(workbook, store),
        // new JvSheet(workbook, store),
        new PostSheet(workbook, store)
    ];


    store.employees.forEach(employee => {
        allSheets.push(new SalesSheet(workbook, employee));
    });

    console.log(store);
}