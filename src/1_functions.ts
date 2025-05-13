function main(workbook: ExcelScript.Workbook) {
    const initialSheets = workbook.getWorksheets();

    const lookup = {};
    const reports = {};

    initialSheets.map(sheet => {
        const sheetName = sheet.getName();
        const data = sheet.getUsedRange().getValues();
        if (sheetName.includes("-")) {
            const splitName = sheetName.split("-");
            if (splitName.length > 1) {
                const typeOfReport = splitName[0].trim().toLowerCase(); // In case we need more control
                const reportName = splitName[1].trim().toLowerCase();
                lookup[reportName] = new LookupReport(data);
            }
        } else {
            switch (sheetName) {
                case "INPUT": reports["globals"] = new InputReport(data);
                    break;
                case "0432": reports["deals"] = new DealsReport(data);
                    break;
                case "90": reports["averageunits"] = new AverageUnitReport(data);
                    break;
                case "3213": reports["priordraw"] = new PriorDrawReport(data);
                    break;
                case "SPIFFS": reports["spiffs"] = new SpiffReport(data);
                    break;
                case "NpsSheet":
                    const head = data[2];
                    const body = data.slice(3)
                    reports["nps"] = new NpsReport([head, ...body]);
                    break;
                default:
                    sheet.delete();
                    break;
            }
        }
    });

    const { storeName, storeAbbr, date, regionalScore, topsalesAmount, ...accounts }: Store = reports["globals"].vars;
    const { retroAcct, expenseAcct1, expenseAcct2, salesTaxAcct, salesBonusAcct1, salesBonusAcct2 }: AccountNumbers = accounts;
    const accountNumbers: AccountNumbers = { retro: retroAcct, expense1: expenseAcct1, expense2: expenseAcct2, salesTax: salesTaxAcct, salesBonus1: salesBonusAcct1, salesBonus2: salesBonusAcct2 }
    const store: Store = {
        name: storeName, abbr: storeAbbr, displayDate: date, employees: [],
        regionalScore: regionalScore, topsalesAmount, accountNumbers
    };

    const employeeList = getAllEmployees(reports);
    attachSpiffs(reports["spiffs"].list, employeeList);
    attachNps(reports["nps"].list, employeeList, store.regionalScore);
    attachPriorDraw(reports["priordraw"].list, employeeList);
    attachUnits(reports["averageunits"].list, employeeList);
    attachDeals(reports["deals"].list, employeeList);
    calculateRetro(lookup["retromini"], lookup["retropercentage"], employeeList);
    calculateUnitBonus(lookup["unit"], employeeList);
    caclulateNpsBonus(regionalScore, employeeList);

    store.employees = employeeList;
    console.log(store);
}


function getAllEmployees(reports: InputReport | DealsReport | AverageUnitReport | PriorDrawReport | SpiffReport | NpsReport) {
    const allEmps: Employee[] = [];
    for (const [key, val] of Object.entries(reports)) {
        switch (key) {
            case "globals":
            case "spiffs":
                break;
            default:
                val.list.map(v => {
                    const employee = allEmps.filter(emp => emp.id === v.employeeID)[0];
                    const nps = new Nps();
                    const averageUnits = new AverageUnits(0);
                    const units = new Units();
                    if (!employee) allEmps.push({ id: v.employeeID, name: v.employeeName, deals: [], nps, spiffsTotal: 0, priorDraw: 0, retroPercent: 0, bonus: { unit: 0, topsales: 0, nps: 0 }, averageUnits, units });
                });
                break;
        }
    }
    return allEmps;
}


function attachSpiffs(spiffs: EmployeeSpiff[], employeeList: Employee[]) {
    spiffs.forEach(spiff => {
        employeeList.filter(emp => emp.id === spiff.employeeID)[0].spiffsTotal = spiff.amount;
    });
}


function attachNps(npsList: EmployeeNps[], employeeList: Employee[], regionalScore: number) {
    npsList.forEach(nps => {
        const employee = employeeList.filter(emp => emp.id === nps.employeeID)[0];
        employee.nps.surveyCount = nps.surveyCount;
        employee.nps.averagePercent = nps.averagePercent;
        employee.nps.currentPercent = nps.currentPercent;
    });
}


function attachPriorDraw(pd: PriorDraw[], employeeList: Employee[]) {
    pd.forEach(d => {
        employeeList.filter(emp => emp.id === d.employeeID)[0].priorDraw = d.draw;
    });
}


function attachUnits(units: EmployeeUnits[], employeeList: Employee[]) {
    units.forEach(unit => {
        const employee = employeeList.filter(emp => emp.id === unit.employeeID)[0];
        employee.averageUnits.threeMonthCount = unit.count;
    });
}


function attachDeals(deals: DealRow[], employeeList: Employee[]) {
    employeeList.forEach(employee => {
        employee.deals = deals.filter(emp => emp.employeeID === employee.id).map(deal => {
            employee.units[deal.vehicleType.toLowerCase()] += deal.dealUnitCount;
            return {
                id: deal.dealID,
                customer: { id: deal.customerID, name: deal.customerName },
                vehicle: new Vehicle(deal.vehicleID, deal.vehicleDescription, deal.vehicleType),
                unitCount: deal.dealUnitCount,
                commission: { fni: deal.commissionFnI, gross: deal.commissionGross, amount: deal.commissionAmount },
                retro: new Retro()
            }
        });
    });
}


function calculateRetro(retroMiniLookup: LookupReport, retroPercentLookup: LookupReport, employeeList: Employee[]) {
    employeeList.forEach(employee => {
        employee.retroPercent = retroPercentLookup.getValue(employee.units.getTotal());
        employee.deals.forEach(deal => {
          const dealRetro = new Retro();
          dealRetro.calculateMini(retroMiniLookup, deal.commission.amount, employee.averageUnits.getAverage(), deal.unitCount);
          dealRetro.calculateOwed(deal.commission.amount);
          dealRetro.calculatePayout(deal.commission.gross, employee.retroPercent);
            deal.retro = dealRetro;
        });
    });
}


function calculateUnitBonus(unitBonusLookup: LookupReport, employeeList: Employee[]) {
    employeeList.forEach(employee => {
        employee.bonus.unit = unitBonusLookup.getValue(employee.units.getTotal());
    });
}

function caclulateNpsBonus(regionalScore: number, employeeList: Employee[]) {
    employeeList.forEach(employee => {
        const unitCount = employee.units.getTotal();
        employee.bonus.nps = employee.nps.getBonus(unitCount, regionalScore);
    });
}