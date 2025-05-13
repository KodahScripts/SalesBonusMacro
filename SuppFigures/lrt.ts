const YEAR = "25";




// Change the YEAR variable and hit 'Run' button.  Book built.
























































enum Color {
    LIGHTBLUE = "#C0E6F5",
    DARKBLUE = "#0F9ED5",
    LIGHTGREY = "#BFBFBF",
    DARKGREY = "#808080",
    YELLOW = "#FFFF00",
}


function main(workbook: ExcelScript.Workbook) {
    const topRows = [
        ["!! Attention !!", '', '', '', '', ''],
        ["Supplemental figures reported as larger than their related accounts on the financial statement may not be used.", '', '', '', '', ''],
        ["(1) DOC/ADMINISTRATIVE FEES INCLUDED IN OTHER INCOME", '', '', '', '', ''],
        ["Any documentary, processing, administrative fee or any hard pack dollars from the sale of new and used vehicles, including new and used F & I and extended service plan income allocated to Account 805. If you produce a separate statement removing these fees and adding those dollars to gross, you do not need to fill out this section.", '', '', '', '', ''],
        ['', '', '', '', "MONTH", "YEAR TO DATE"],
        ["New Jaguar Vehicles", '', '', '', '', "=E6"],
        ["New Land Rover Vehicles", '', '', '', '', "=E7"],
        ["New Other Vehicles", '', '', '', '', "=E8"],
        ["New F&I", '', '', '', '', "=E9"],
        ["Used CPO Vehicles", '', '', '', '', "=E10"],
        ["Used Non-CPO Vehicles", '', '', '', '', "=E11"],
        ["Used F&I", '', '', '', '', "=E12"],
        ["TOTAL (MUST BE LESS THAN OR EQUAL TO ACCOUNTS 9804, 9805 & 9890 OF YOUR FINANCIAL STATEMENT)", '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ["(2) LEASE TERMINATION INFORMATION", '', '', '', '', ''],
        ["Office Manager and F&I Manager to set up a procedure for tracking this information.  New terminations = vehicles originally leased when new.  Used terminations = vehicles originally leased when used (includes program cars).", '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", ''],
        ['', '', "NEW", "USED", "NEW", "USED"],
        ["A. # Leases Renewed (We Leased the Customer Another Vehicle)", '', '', '', "=C20", "=D20"],
        ["B. # Lease Customers Who Returned the Vehicle & Walked Away", '', '', '', "=C21", "=D21"],
        ["C. # Customers Who Purchased Lease Return Unit or Another Vehicle from Dealership", '', '', '', "=C22", "=D22"],
        ["D. TOTAL LEASE TERMINATIONS (A+B+C)", '', "=C20+C21+C22", "=D20+D21+D22", "=E20+E21+E22", "=F20+F21+F22"],
        ['', '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ["(3) VEHICLES FINANCED", '', '', '', '', ''],
        ["Retail lease = any contract not titled to the retail customer.", '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", ''],
        ['', '', "NEW", "USED", "NEW", "USED"],
        ["# Retail Vehicles Leased (Leased ONLY)", '', '', '', "=C30", "=D30"],
        ["All Chargebacks of Finance Income (Excluding Insurance Chargebacks)", '', '', '', "=C31", "=D31"],
        ['', '', '', '', '', ''],
        ["(5) GAP POLICIES", '', '', '', '', ''],
        ["Gross income is the dealership gross/commission from the GAP insurance contract.(Do not include physical damage or liability commissions.)", '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", ''],
        ['', '', "NEW", "USED", "NEW", "USED"],
        ["All Chargebacks of GAP Insurance Income", '', '', '', "=C37", "=D37"],
        ['', '', '', '', '', ''],
        ["(6) EXTENDED SERVICE CONTRACTS", '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", ''],
        ['', '', "NEW", "USED", "NEW", "USED"],
        ["All Chargebacks of Service Contract Income", '', '', '', "=C42", "=D42"],
        ['', '', '', '', '', ''],
        ["(7) MAINTENANCE CONTRACTS", '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", ''],
        ['', '', "NEW", "USED", "NEW", "USED"],
        ["All Chargebacks of Maintenance Contract Income", '', '', '', "=C47", "=D47"],
        ['', '', '', '', '', ''],
        ["(8) OTHER CONTRACTED PROTECTION PRODUCTS (Tire & Rim, Key Replacement, Etch, Lease End Wear Care, etc.)", '', '', '', '', ''],
        ["Any gross profit on the following contract term products (Tire & Rim, Key Replacement, Etch, Lease End Wear Care, etc.) sold by the F&I department (and not included in Vehicle Gross).", '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", ''],
        ['', '', "NEW", "USED", "NEW", "USED"],
        [" GP Other Contracted Protection Products Sold by F&I", '', '', '', "=C53", "=D53"],
        ['', '', '', '', '', ''],
        ["(9) OTHER NON-CONTRACTED MERCHANDISE/AFTERMARKET (Accessories, Paint/Fabric/Rust, 3M, etc.)", '', '', '', '', ''],
        ["Any gross profit on accessories or other merchandise that are not contract term based (Paint/Fabric/Rust, 3M, etc.) sold by the F&I department (and not included in Vehicle Gross).", '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", ''],
        ['', '', "NEW", "USED", "NEW", "USED"],
        [" GP Other Non-Contracted", '', '', '', "=C59", "=D59"],
        ['', '', '', '', '', ''],
        ["(12) SERVICE LABOR ANALYSIS", '', '', '', '', ''],
        ["We will check your Hours reported in each labor category by dividing the Sales $$ on the statement by the Hours Reported. If the resulting calculated Effective Labor Rate is too high (> $200) or too low (< $25), the Hours provided WILL NOT be used. We will also automatically update any obvious decimal related errors.", '', '', '', '', ''],
        ["LAND ROVER - SERVICE LABOR ANALYSIS", "POSTED LABOR RATE", "EFFECTIVE LABOR RATE - MONTH", "EFFECTIVE LABOR RATE - YEAR TO DATE", "HOURS SOLD - MONTH", "HOURS SOLD - YEAR TO DATE"],
        ["Customer", '', '', '', '', "=E64"],
        ["Warranty Claim", '', '', '', '', "=E65"],
        ["Internal", '', '', '', '', "=E66"],
        ["Factory Paid Scheduled Maintenance", '', '', '', '', "=E67"],
        ["TOTAL HOURS SOLD", '', '', '', "=E64+E65+E66+E67", "=F64+F65+F66+F67"],
        ['', '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ["(27) INTERNET SALES INFORMATION - MONTH & YTD", '', '', '', '', ''],
        ["MONTH", '', '', '', '', ''],
        ['', '', '', '', "NEW VEHICLES", "USED VEHICLES"],
        ["Your Website Leads", '', '', '', '', ''],
        ["Manufacturers Leads", '', '', '', '', ''],
        ["Lead Provider Leads", '', '', '', '', ''],
        ["TOTAL LEADS", '', '', '', "=E74+E75+E76", "=F74+F75+F76"],
        ['', '', '', '', '', ''],
        ["# Internet Sales", '', '', '', '', ''],
        ["Total Internet Front Gross", '', '', '', '', ''],
        ["Total Internet F&I Gross", '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ["(27) INTERNET SALES INFORMATION - MONTH & YTD", '', '', '', '', ''],
        ["YEAR TO DATE", '', '', '', '', ''],
        ['', '', '', '', "NEW VEHICLES", "USED VEHICLES"],
        ["Your Website Leads", '', '', '', "=E74", "=F74"],
        ["Manufacturers Leads", '', '', '', "=E75", "=F75"],
        ["Lead Provider Leads", '', '', '', "=E76", "=F76"],
        ["TOTAL LEADS", '', '', '', "=E86+E87+E88", "=F86+F87+F88"],
        ['', '', '', '', '', ''],
        ["# Internet Sales", '', '', '', "=E79", "=F79"],
        ["Total Internet Front Gross", '', '', '', "=E80", "=F80"],
        ["Total Internet F&I Gross", '', '', '', "=E81", "=F81"]
    ];

    const tabs = ["01.", "02.", "03.", "04.", "05.", "06.", "07.", "08.", "09.", "10.", "11.", "12."];
    const formulaRows = [5,6,7,8,9,10,11,12,19,20,21,29,30,36,41,46,52,58,63,64,65,66,85,86,87,90,91,92];
    let lastSheetName = "";
    tabs.forEach(tab => {
        const sheetName = tab + YEAR;
        const sheet = workbook.addWorksheet(sheetName);
        const month = Number(tab.split('.')[0]);
        switch (month) {
            case 1:
                sheet.getRangeByIndexes(0, 0, topRows.length, topRows[0].length).setValues(topRows);
                break;
            default:
                let rows = topRows.map((row, index) => {
                    return row.map((r, i) => {
                        if (index > 82) {
                            if (r.startsWith("=") && formulaRows.includes(index)) {
                                const cell = r.slice(1);
                                const cellNumber = Number(cell.slice(1)) + 12;
                                const othercell = i === 4 ? `E${cellNumber}` : `F${cellNumber}`;
                                return `=${cell} + '${lastSheetName}'!${othercell}`;
                            }
                        } else {
                            if (r.startsWith("=") && formulaRows.includes(index)) {
                                const cell = r.slice(1);
                                const cellNumber = cell.slice(1);
                                const othercell = i === 4 ? `E${cellNumber}` : `F${cellNumber}`;
                                return `=${cell} + '${lastSheetName}'!${othercell}`;
                            }
                        }
                        return r;
                    });
                });
                sheet.getRangeByIndexes(0, 0, topRows.length, topRows[0].length).setValues(rows);
                break;
        }
        lastSheetName = sheetName;
        stylize(sheet);
    });
}


function stylize(sheet: ExcelScript.Worksheet) {
    const rows = sheet.getUsedRange().getValues();
    const row1 = sheet.getRange("1:1").getFormat();
    const colA = sheet.getRange("A:A").getFormat();

    const full = [1,2,3,4,14,15,16,17,24,25,26,27,32,33,34,38,39,43,44,48,49,50,54,55,56,60,61,62,69,70,71,72,78,82,83,84,90];
    const mid = [5,6,7,8,9,10,11,12,13,73,74,75,76,77,79,80,81,85,86,87,88,89,91,92,93];
    const small = [18,19,20,21,22,23,28,29,30,31,35,36,37,40,41,42,45,46,47,51,52,53,57,58,59,];
    const doubleMerge = [18,28,35,40,45,51,57];
    const bldCntr = [5,19,29,36,41,46,52,58,63,73,85];
    const blue = [1,2,3,4,16,17,26,27,33,34,39,44,49,50,55,56,61,62,63,71,72,83,84];
    const grey = ["A8:F8", "A11:F11", "A20:F22", "A30:F31", "A8:F8", "A37:F37", "A42:F42", "A47:F47", "A53:F53", "A59:F59", "A64:F67", "E74:F76", "E79:F81", "E86:F88", "E91:F93"];

    colA.getFont().setBold(true);
    rows.map((row, index) => {
        const currentRow = index + 1;

        if (blue.includes(currentRow)) {
            const rng = sheet.getRange(`A${currentRow}`).getFormat();
            rng.getFill().setColor(Color.LIGHTBLUE);
            rng.getFont().setBold(true);
            rng.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
        }
        if (doubleMerge.includes(currentRow)) {
            const rng1 = sheet.getRange(`C${currentRow}:D${currentRow}`);
            const rng2 = sheet.getRange(`E${currentRow}:F${currentRow}`);
            rng1.merge();
            rng2.merge();
            rng1.getFormat().getFont().setBold(true);
            rng2.getFormat().getFont().setBold(true);
            rng1.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
            rng2.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        }
        if (full.includes(currentRow)) sheet.getRange(`A${currentRow}:F${currentRow}`).merge();
        if (mid.includes(currentRow)) sheet.getRange(`A${currentRow}:D${currentRow}`).merge();
        if (small.includes(currentRow)) sheet.getRange(`A${currentRow}:B${currentRow}`).merge();
        if (bldCntr.includes(currentRow)) {
            const rng = sheet.getRange(`A${currentRow}:F${currentRow}`).getFormat();
            rng.getFont().setBold(true);
            rng.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        }
    });

    grey.forEach(g => {
        sheet.getRange(g).getFormat().getFill().setColor(Color.LIGHTGREY);
    });


    [sheet.getRange("A1:F13").getFormat(),
    sheet.getRange("A16:F23").getFormat(),
    sheet.getRange("A26:F31").getFormat(),
    sheet.getRange("A33:F37").getFormat(),
    sheet.getRange("A39:F42").getFormat(),
    sheet.getRange("A44:F47").getFormat(),
    sheet.getRange("A49:F53").getFormat(),
    sheet.getRange("A55:F59").getFormat(),
    sheet.getRange("A61:F68").getFormat(),
    sheet.getRange("A71:F81").getFormat(),
    sheet.getRange("A83:F93").getFormat()].forEach(borderRange => {
        borderRange.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
        borderRange.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin);
        borderRange.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin);
        borderRange.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin);
        borderRange.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin);
        borderRange.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(ExcelScript.BorderWeight.thin);
    });


    sheet.getUsedRange(true).getFormat().setWrapText(true);
    row1.setColumnWidth(100);
    colA.setRowHeight(100);
    colA.autofitRows();
    row1.autofitColumns();
    row1.getFont().setColor('red');
    row1.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
}