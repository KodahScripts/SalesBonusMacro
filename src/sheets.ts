class Sheet {
    protected sheet: ExcelScript.Worksheet;
    constructor(protected workbook: ExcelScript.Workbook, protected sheetName: string, protected columnNames: Array<string>, protected headerRow: number = 0) {
        this.sheet = workbook.addWorksheet(sheetName);
        this.sheet.getRangeByIndexes(headerRow, 0, 1, columnNames.length).setValues([columnNames]);
    }
}

class DisplayDate {
    public descDate: string;
    public today: number;
    constructor() {
        const date = new Date().toISOString().slice(0, 10);
        const todayArr = date.split('-');
        const descYear = todayArr[0].split('0')[1];
        this.descDate = [todayArr[1], descYear].join('');
        this.today = Number([todayArr[1], todayArr[2], descYear].join(''));     
    }
}

class NpsSheet extends Sheet {
    constructor(protected workbook: ExcelScript.Workbook, private store: Store) {
        super(workbook, "NPS", ["Employee #", "Salesperson", "# of Surveys", "Current Score", "90 Day Score", "Score for Bonus", "CSI Outcome", "Regional Score", String(store.regionalScore)]);
        store.employees.map((employee, index) => {
            const row = index + 1;
            const data = [employee.id, employee.name, employee.nps.surveys, employee.nps.current / 100, employee.nps.average / 100, `=IF(D${row}>E${row},D${row},E${row})`, employee.nps.outcome];
            this.sheet.getRangeByIndexes(row, 0, 1, data.length).setValues([data]);
        });
        this.sheet.addTable(`A1:G${store.employees.length + 1}`, true);
        this.sheet.getRange(`D2:G${store.employees.length + 1}`).setNumberFormat("0.0%");
        this.sheet.getRange("H1:I1").getFormat().getFont().setBold(true);
        this.sheet.getRange("1:1").getFormat().autofitColumns();
    }
}

class PaySummarySheet extends Sheet {
    constructor(protected workbook: ExcelScript.Workbook, private store: Store) {
        super(workbook, "Pay Summary", ["Employee #", "Salesperson", "Total Units", "Gross for Rank", "Bonus Rank", "F&I Totals", "Spiffs to Pay", "Commission 3120", "Retro Commission", "F&I Commission", "Month End Bonus 3122", "Total EOM Bonus 8328", "Draw 3121", "Total Commissions", "YTD Bucket", "Deposit Gross", "Check - Should be 0", "Draw to Take", "Expense 1", "Expense 2"]);
        store.employees.map((employee, index) => {
            const row = index + 1;
            const rowStr = index + 2;
            const rankFormula = `=IF(C${rowStr}<15,"",RANK.EQ(C${rowStr},C:C,0))`;
            const depositGross = employee.getDepositGross();
            const balanceComm = `=H${rowStr} + ${employee.commissionBalance}`;
            const data = [employee.id, employee.name, employee.units.total, employee.commission.gross.toFixed(2), rankFormula, employee.commission.fni.toFixed(2), employee.spiff.toFixed(2), employee.commission.amount.toFixed(2), employee.retro.total.toFixed(2), employee.fni.payout.toFixed(2), employee.bonus.total.toFixed(2), employee.bonus.eom.toFixed(2), employee.priorDraw.toFixed(2), employee.totalCommission.toFixed(2), employee.ytdBucket.toFixed(2), depositGross.toFixed(2), balanceComm, employee.drawAmount.toFixed(2), employee.expense.one.toFixed(2), employee.expense.two.toFixed(2)];
            this.sheet.getRangeByIndexes(row, 0, 1, data.length).setValues([data]);
        });
        const unitRange = this.sheet.getRangeByIndexes(store.employees.length + 4, 2, 3, 3);
        unitRange.setValues([
            ["UNITS", store.units.total, ''],
            ["NEW", store.units.new, store.unitPercent.new.toFixed(2)],
            ["USED", store.units.used, store.unitPercent.used.toFixed(2)]
        ]);
        unitRange.getFormat().getFont().setBold(true);
        const totalData = [
            [store.commission.taxes.toFixed(2), "SLSTAX"],
            [(store.commission.taxes * store.unitPercent.new).toFixed(2), "SLSBNSTAX"],
            [(store.commission.taxes * store.unitPercent.used).toFixed(2), "SLSBNSTAX"],
            [store.commission.fni, "TOTAL BONUS"]
        ];
        const totalsDataRange = this.sheet.getRangeByIndexes(store.employees.length + 4, 6, 4, 2);
        const percentRange = this.sheet.getRangeByIndexes(store.employees.length + 4, 4, 4, 1);
        percentRange.setNumberFormat("0%");
        totalsDataRange.setValues(totalData);
        totalsDataRange.getFormat().getFont().setBold(true);
        totalsDataRange.setNumberFormat(NumberFormat.ACCOUNTING);
        this.sheet.addTable(`A1:T${store.employees.length + 1}`, true).setShowTotals(true);
        this.sheet.getRange(`D2:D${store.employees.length + 1}`).setNumberFormat(NumberFormat.CURRENCY);
        this.sheet.getRange(`F2:T${store.employees.length + 1}`).setNumberFormat(NumberFormat.CURRENCY);
        this.sheet.getRange("1:1").getFormat().autofitColumns();
    }
}

class PostSheet extends Sheet {
    constructor(protected workbook: ExcelScript.Workbook, private store: Store) {
        super(workbook, "PowerPost", ["DATE", "GL", "AMOUNT", "EMPID", "DESC"]);
        const { descDate, today } = new DisplayDate();
        const dataArr: Array<string | number>[] = [];
        store.employees.map(employee => {
            const remainder = employee.getRemainder().toFixed(2);
            const expenseOne = employee.expense.one.toFixed(2);
            const expenseTwo = employee.expense.two.toFixed(2);
            dataArr.push(
                [`=TEXT(${today}, "000000")`, store.accounts.retro, Number(remainder) * -100, employee.id, `SLSBONUS${descDate}`],
                [`=TEXT(${today}, "000000")`, store.accounts.expense.one, Number(expenseOne) * 100, employee.id, `EXPENSE1${descDate}`],
                [`=TEXT(${today}, "000000")`, store.accounts.expense.two, Number(expenseTwo) * 100, employee.id, `EXPENSE2${descDate}`]
            );
        });
        dataArr.push([`=TEXT(${today}, "000000")`, store.accounts.salesTax, store.commission.gross * 100, '', `SLSTAX${descDate}`]);
        dataArr.push([`=TEXT(${today}, "000000")`, store.accounts.salesBonusTax1, (store.commission.gross * store.unitPercent.new) * -100, '', `SLSBNSTAX${descDate}`]);
        dataArr.push([`=TEXT(${today}, "000000")`, store.accounts.salesBonusTax2, (store.commission.gross * store.unitPercent.used) * -100, '', `SLSBNSTAX${descDate}`]);
        dataArr.map((data, index) => {
            const row = index + 1;
            this.sheet.getRangeByIndexes(row, 0, 1, data.length).setValues([data]);
        });
    }
}

class SalesSheet extends Sheet {
    constructor(protected workbook: ExcelScript.Workbook, private employee: Employee) {
        super(workbook, employee.name, ["Date", "Reference #", "Customer #", "Customer Name", "Stock #", "Year", "Make", "Model", "Sale Type", "Commission F&I", "Commission Gross", "Units", "Commission Amount", "Retro Mini", "Retro Owed", "Retro Commission Payout"], 7);
        const headerData = [
            ["Name", employee.name],
            ["Employee Number", employee.id],
            ["90 Day Rolling Average #", employee.averageUnits],
            ["CSI", employee.nps.outcome],
            ["# of Surveys", employee.nps.surveys],
            ["Retro Percentage", employee.getRetroPercentage()]
        ];
        this.sheet.getRangeByIndexes(0, 0, headerData.length, headerData[0].length).setValues(headerData);
        const dataStartRow = this.headerRow + 1;
        let dealData: Array<string | number>[] = [];
        if (employee.deals.length > 0) {
            dealData = employee.deals.map(deal => {
                return [deal.date, deal.id, deal.customer.id, deal.customer.name, deal.vehicle.id, deal.vehicle.year, deal.vehicle.make, deal.vehicle.model, deal.vehicle.saleType, deal.commission.fni.toFixed(2), deal.commission.gross.toFixed(2), deal.unitCount, deal.commission.amount.toFixed(2), deal.retro.mini.toFixed(2), deal.retro.owed.toFixed(2), deal.retro.payout.toFixed(2)];
            });
            this.sheet.getRangeByIndexes(dataStartRow, 0, dealData.length, dealData[0].length).setValues(dealData);
        }
        const totalsData = [
            ["Prior Draw Balance", '', employee.priorDraw],
            ["Commission", 0.18, employee.commission.amount],
            ["Retro Commission", employee.getRetroPercentage(), employee.retro.payout],
            ["Retro Mini", '', employee.retro.owed],
            ["Total Retro Commission", '', employee.retro.total],
            ["Total F&I", '', employee.commission.fni],
            ["25% Reserve F&I", -0.25, employee.fni.reserve],
            ["Total F&I Payable Gross", '', employee.fni.gross],
            ["Total F&I Payout", 0.05, employee.fni.payout],
            ["Top Salesman Bonus", '', employee.bonus.topsales],
            ["Unit Bonus", employee.units.total, employee.bonus.unit],
            ["CSI", employee.nps.outcome, employee.bonus.csi],
            ["Total Bonus", '', employee.bonus.total],
            ["Spiff", '', employee.spiff],
            ["Total Pay", '', employee.getTotalPay()],
            ["Bucket Total YTD", '', employee.calcNewBucket()]
        ];
        const totalsStartRow = dataStartRow + dealData.length + 2;
        this.sheet.getRangeByIndexes(totalsStartRow, 0, totalsData.length, totalsData[0].length).setValues(totalsData);
    }
}