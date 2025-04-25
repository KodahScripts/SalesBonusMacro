class SalesSheet {
    private sheet: ExcelScript.Worksheet;
    constructor(protected workbook: ExcelScript.Workbook, protected employee: Employee) {
        if (employee.deals.length === 0) return;
        this.sheet = workbook.addWorksheet(employee.name);
        const startRow = 8;
        const dealCount = employee.deals.length;
        const lastRow = dealCount + startRow;
        const headerRange = this.sheet.getRange("A1:B6");
        const colHeaderRange = this.sheet.getRange("A7:P7");

        headerRange.setValues([
            ["Name", employee.name],
            ["Employee Number", employee.id],
            ["90 Day Rolling Average #", employee.averageUnits],
            ["CSI", employee.nps.outcome],
            ["# of Surveys", employee.nps.surveys],
            ["Retro Percentage", employee.getRetroPercentage()]
        ]);

        colHeaderRange.setValues([["Date", "Reference #", "Customer #", "Customer Name", "Stock #", "Year", "Make", "Model", "Sale Type", "Commission F&I", "Commission Gross", "Units", "Commission Amount", "Retro Mini", "Retro Owed", "Retro Commission Payout"]]);

        employee.deals.forEach((deal, index) => {
            const row = index + 8;
            this.sheet.getRange(`A${row}:P${row}`).setValues([[
                deal.date,
                deal.id,
                deal.customer.id,
                deal.customer.name,
                deal.vehicle.id,
                deal.vehicle.year,
                deal.vehicle.make,
                deal.vehicle.model,
                deal.vehicle.saleType,
                deal.commission.fni.toFixed(2),
                deal.commission.gross.toFixed(2),
                deal.unitCount,
                deal.commission.amount.toFixed(2),
                deal.retro.mini.toFixed(2),
                deal.retro.owed.toFixed(2),
                deal.retro.payout.toFixed(2)
            ]]);
        });

        this.sheet.getRange(`J${lastRow}:P${lastRow}`).setValues([[
            employee.commission.fni.toFixed(2),
            employee.commission.gross.toFixed(2),
            employee.units.total,
            employee.commission.amount.toFixed(2),
            employee.retro.mini.toFixed(2),
            employee.retro.owed.toFixed(2),
            employee.retro.payout.toFixed(2)
        ]]);

        const totalsStartRow = lastRow + 2;
        const employeeSignatureRow = totalsStartRow + 3;
        const managerSignatureRow = employeeSignatureRow + 6;
        const employeeSignatureRange = this.sheet.getRange(`B${employeeSignatureRow}:D${employeeSignatureRow}`);
        const managerSignatureRange = this.sheet.getRange(`B${managerSignatureRow}:D${managerSignatureRow}`);
        this.sheet.getRange(`E${employeeSignatureRow}`).setValue("EMPLOYEE");
        this.sheet.getRange(`E${managerSignatureRow}`).setValue("MANAGER");

        const totalsLabels = [
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
        totalsLabels.forEach((label, index) => {
            const row = totalsStartRow + index;
            this.sheet.getRange(`I${row}:K${row}`).setValues([label]);
        });

        employeeSignatureRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
        managerSignatureRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);

        this.format();
    }

    format() {
        const dealStartRow = 8;
        const dealCount = this.employee.deals.length;
        const reportRange = this.sheet.getRange(`A7:P${dealCount + dealStartRow}`);
        const reportTable = this.sheet.addTable(reportRange, true);
        reportTable.setPredefinedTableStyle("TableStyleLight2");
        reportTable.setShowFilterButton(false);
        this.sheet.getRange(`A${dealStartRow}:A${dealCount + dealStartRow}`).setNumberFormat(NumberFormat.DATE);
        
        this.sheet.getRange("A:C").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        this.sheet.getRange("A:C").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
        this.sheet.getRange("E:F").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        this.sheet.getRange("E:F").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
        this.sheet.getRange("I:P").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        this.sheet.getRange("I:P").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
        reportRange.getFormat().autofitColumns();
    }
}