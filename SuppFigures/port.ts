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
        ["PORSCHE TAMPA- 8318-56", '', '', '', '', '', '', ''],
        ["(1) DOC/ADMINISTRATIVE FEES INCLUDED IN OTHER INCOME", '', '', '', '', '', '', ''],
        ["Any documentary, processing, administrative fee or any hard pack dollars from the sale of new and used vehicles, including new and used F & I and extended service plan income allocated to Account 805. If you produce a separate statement removing these fees and adding those dollars to gross, you do not need to fill out this section.", '', '', '', '', '', "PULL 0504 USE CURRENT BALANCE", ''],
        ['', '', '', '', "MONTH", "YEAR TO DATE", '', ''],
        ["New Porsche Vehicles", '', '', '', '', "=E5", '5810A', ''],
        ["New Other Vehicles", '', '', '', '', '', '', ''],
        ["New F&I", '', '', '', '', '', '', ''],
        ["Used CPO Vehicles", '', '', '', '', "=E8", '5810B', ''],
        ["Used Non-CPO Vehicles", '', '', '', '', '', '', ''],
        ["Used Other Vehicles", '', '', '', '', '', '', ''],
        ["Used F&I", '', '', '', '', '', '', ''],
        ["TOTAL (MUST BE LESS THAN OR EQUAL TO ACCOUNTS 9151/9351 & 9191/9391 OF YOUR FINANCIAL STATEMENT)", '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', ''],
        ["(2) LEASE TERMINATION INFORMATION", '', '', '', '', '', '', ''],
        ["Office Manager and F&I Manager to set up a procedure for tracking this information.  New terminations = vehicles originally leased when new.  Used terminations = vehicles originally leased when used (includes program cars).", '', '', '', '', '', 'FROM STORE', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", '', '', ''],
        ['', '', "NEW", "USED", "NEW", "USED", '', ''],
        ["A. # Leases Renewed (We Leased the Customer Another Vehicle)", '', '', '', "=C18", "=D18", '', ''],
        ["B. # Lease Customers Who Returned the Vehicle & Walked Away", '', '', '', "=C19", "=D19", '', ''],
        ["C. # Customers Who Purchased Lease Return Unit or Another Vehicle from Dealership", '', '', '', "=C20", "=D20", '', ''],
        ["D. TOTAL LEASE TERMINATIONS (A+B+C)", '', "=C18+C19+C20", "=D18+D19+D20", "=E18+E19+E20", "=F18+F19+F20", '', ''],
        ['', '', '', '', '', '', '', ''],
        ["(3) VEHICLES FINANCED", '', '', '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", '', '', ''],
        ['', '', "NEW", "USED", "NEW", "USED", '', ''],
        ["# Vehicles Financed (Excluding Leases)", '', '', '', "=C26", "=D26", "DOC VISION- DLRSUMMARY REPORT - Financed/ Lease line, retail column", ''],
        ["# Retail Vehicles Leased", '', '', '', "=C27", "=D27", "DOC VISION- DLRSUMMARY REPORT - Financed/ Lease line , Lease column", ''],
        ["Total Finance/Lease Reserve Gross Income (Excluding Insurance)", '', '', '', "=C28", "=D28", "DOC VISION- DLRSUMMARY REPORT - Financed/ Lease line, retail column", ''],
        ["All Chargebacks of Finance Income(Excluding Insurance Chargebacks)", '', '', '', "=C29", "=D29", "MAGHON Stmt - pg2 , line 8, 1st column. Click thru - accts 5772, 5782", ''],
        ['', '', '', '', '', '', '', ''],
        ["(5) GAP POLICIES", '', '', '', '', '', '', ''],
        ["Gross income is the dealership gross/commission from the GAP insurance contract.(Do not include physical damage or liability commissions.)", '', '', '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", '', '', ''],
        ['', '', "NEW", "USED", "NEW", "USED", '', ''],
        ["# GAP Policies Sold", '', '', '', "=C35", "=D35", "FI Mgr Summary in DocVision", ''],
        ["GAP Gross Income", '', '', '', "=C36", "=D36", "FI Mgr Summary in DocVision", ''],
        ["All Chargebacks of GAP Insurance Income", '', '', '', "=C37", "=D37", "MAGHON Stmt - pg2 , line 8, 1st column. Click thru - Gap accts", ''],
        ['', '', '', '', '', '', '', ''],
        ["(6) EXTENDED SERVICE CONTRACTS", '', '', '', '', '', '', ''],
        ["Gross income amount should reflect only the amount of gross profit (not the sale price). [If a dealer owns his company, cost of program administration and reserve amounts should NOT be included in gross profit.]", '', '', '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", '', '', ''],
        ['', '', "NEW", "USED", "NEW", "USED", '', ''],
        ["# Extended Service Contracts (Excludes Maintenance Contracts)", '', '', '', "=C43", "=D43", "FI Mgr Summary in DocVision", ''],
        ["Extended Service Contract Gross Income", '', '', '', "=C44", "=D44", "FI Mgr Summary in DocVision", ''],
        ["All Chargebacks of Service Contract Income", '', '', '', "=C45", "=D45", "MAGHON Stmt - pg2 , line 8, 1st column. Click thru - svc K accts", ''],
        ['', '', '', '', '', '', '', ''],
        ["(7) MAINTENANCE CONTRACTS", '', '', '', '', '', '', ''],
        ["If a customer is sold any maintenance contract, only one 'vehicle contract' should be shown. Gross income amount should reflect only the amount of gross profit (not the sale price). [If a dealer owns his company, cost of program administration and reserve amounts should NOT be included in gross profit.]", '', '', '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", '', '', ''],
        ['', '', "NEW", "USED", "NEW", "USED", '', ''],
        ["# Maintenance Contracts", '', '', '', "=C51", "=D51", "FI Mgr Summary in DocVision", ''],
        ["Maintenance Contract Gross Income", '', '', '', "=C52", "=D52", "FI Mgr Summary in DocVision", ''],
        ["All Chargebacks of Maintenance Contract Income", '', '', '', "=C53", "=D53", "MAGHON Stmt - pg2 , line 8, 1st column. Click thru - car care accts PPD", ''],
        ['', '', '', '', '', '', '', ''],
        ["(8) OTHER CONTRACTED PROTECTION PRODUCTS (Tire & Rim, Key Replacement, Etch, Lease End Wear Care, etc.)", '', '', '', '', '', '', ''],
        ["Any gross profit on the following contract term products (Tire & Rim, Key Replacement, Etch, Lease End Wear Care, etc.) sold by the F&I department (and not included in Vehicle Gross).", '', '', '', '', '', "FI Mgr Summary in DocVision - Haz, Etch, Dent, Wear, Key, Windshield, UVP, Tire", ''],
        ['', '', "MONTH", '', "YEAR TO DATE", '', '', ''],
        ['', '', "NEW", "USED", "NEW", "USED", '', ''],
        [" GP Other Contracted Protection Products Sold by F&I", '', '', '', "=C59", "=D59", '', ''],
        ['', '', '', '', '', '', '', ''],
        ["(9) OTHER NON-CONTRACTED MERCHANDISE/AFTERMARKET (Accessories, Paint/Fabric/Rust, 3M, etc.)", '', '', '', '', '', '', ''],
        ["Any gross profit on accessories or other merchandise that are not contract term based (Paint/Fabric/Rust, 3M, etc.) sold by the F&I department (and not included in Vehicle Gross).", '', '', '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", '', '', ''],
        ['', '', "NEW", "USED", "NEW", "USED", '', ''],
        [" GP Other Non-Contracted Accessories/Other Merchandise (Aftermarket) Sold by F&I", '', '', '', "=C65", "=D65", '', ''],
        ['', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', ''],
        ["(11) SERVICE LABOR ANALYSIS", '', '', '', '', '', '', ''],
        ["Porsche - SERVICE LABOR ANALYSIS", "POSTED LABOR RATE", "EFFECTIVE LABOR RATE - MONTH", "EFFECTIVE LABOR RATE - YEAR TO DATE", "HOURS SOLD - MONTH", "HOURS SOLD - YEAR TO DATE", '', ''],
        ["Customer", '', '', '', '', "=E70", '', ''],
        ["Warranty Claim", '', '', '', '', "=E71", '', ''],
        ["Internal", '', '', '', '', "=E72", '', ''],
        ["TOTAL HOURS SOLD", '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', ''],
        ["(16) CURRENT FLOOR PLAN INTEREST RATES", '', '', '', '', "BOA-REPORTING-REPORTS-MNTHLY INT RATE", '', ''],
        ['', '', '', '', '', '', '', ''],
        ["(23) NEW INVENTORY", '', '', '', "CAR", "TRUCK", '', ''],
        ["New Vehicles in Inventory", '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', ''],
        ["(27)NEW INVENTORY", '', '', '', "1-30 DAYS", "31-60 DAYS", "61-90 DAYS", "91+"],
        ["NEW UNIT COUNT", '', '', '', '', '', '', ''],
        ["NEW INVENTORY AMOUNT", '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', ''],
        ["(27)USED INVENTORY", '', '', '', "1-30 DAYS", "31-60 DAYS", "61-90 DAYS", "91+"],
        ["USED UNIT COUNT", '', '', '', '', '', '', ''],
        ["USED INVENTORY AMOUNT", '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', ''],
        ["(28) INTERNET SALES INFORMATION - MONTH & YTD", '', '', '', '', '', '', ''],
        ["MONTH", '', '', '', '', '', '', ''],
        ['', '', '', '', "NEW VEHICLES", "USED VEHICLES", '', ''],
        ["Your Website Leads", '', '', '', '', '', '', ''],
        ["Manufacturers Leads", '', '', '', '', '', '', ''],
        ["Lead Provider Leads", '', '', '', '', '', '', ''],
        ["TOTAL LEADS", '', '', '', "=E92+E93+E94", "=F92+F93+F94", '', ''],
        ['', '', '', '', '', '', '', ''],
        ["# Internet Sales", '', '', '', '', '', '', ''],
        ["Total Internet Front Gross", '', '', '', '', '', '', ''],
        ["Total Internet F&I Gross", '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', ''],
        ["(28) INTERNET SALES INFORMATION - MONTH & YTD", '', '', '', '', '', '', ''],
        ["YEAR TO DATE", '', '', '', '', '', '', ''],
        ['', '', '', '', "NEW VEHICLES", "USED VEHICLES", '', ''],
        ["Your Website Leads", '', '', '', "=E92", "=F92", '', ''],
        ["Manufacturers Leads", '', '', '', "=E93", "=F93", '', ''],
        ["Lead Provider Leads", '', '', '', "=E94", "=F94", '', ''],
        ["TOTAL LEADS", '', '', '', "=E104+E105+E106", "=F104+F105+F106", '', ''],
        ['', '', '', '', '', '', '', ''],
        ["# Internet Sales", '', '', '', "=E97", "=F97", '', ''],
        ["Total Internet Front Gross", '', '', '', "=E98", "=F98", '', ''],
        ["Total Internet F&I Gross", '', '', '', "=E99", "=F99", '', '']
    ];

    const tabs = ["01.", "02.", "03.", "04.", "05.", "06.", "07.", "08.", "09.", "10.", "11.", "12."];
    const formulaRows = [4, 7, 17, 18, 19, 25, 26, 27, 28, 34, 35, 36, 42, 43, 44, 50, 51, 52, 59, 64, 69, 70, 71, 102,103, 104, 105, 108, 109, 110];
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
                        if (index > 100) {
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

    const full = [1, 2, 3, 14, 15, 23, 31, 32, 39, 40, 47, 48, 55, 56, 61, 62, 68, 89, 90, 96, 100, 101, 102];
    const mid = [4, 5, 6, 7, 8, 9, 10, 11, 12, 76, 78, 79, 81, 82, 83, 85, 86, 87, 91, 92, 93, 94, 95, 97, 98, 99, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111];
    const small = [16, 17, 18, 19, 20, 21, 24, 25, 26, 27, 28, 29, 33, 34, 35, 36, 37, 41, 42, 43, 44, 45, 49, 50, 51, 52, 53, 57, 58, 59, 63, 64, 65, ];
    const doubleMerge = [16, 24, 33, 41, 49, 57, 63];
    const yellow = [1, 2, 3, 14, 15, 23, 24, 25, 31, 32, 33, 39, 40, 41, 47, 48, 49, 55, 56, 61, 62, 68, 76, 78, 81, 85, 89, 90, 101, 102];
    const bldCntr = [4, 17, 25, 34, 42, 50, 58, 64, 69, 78, 81, 85, 91, 103];
    const grey = [12, 21, 73, 95, 107];

    const lblue = ["A5:E5", "A8:E8", "C18:D20", "C26:D29", "C35:D37", "C43:D45", "C51:D53", "C59:D59", "C65:D65", "B69:F69", "B70:E72", "E76", "E79:F79", "E82:H83", "E86:H87", "E92:F94", "E97:F99", "E104:F106", "E109:F111"];
    const dblue = ["F5", "F8", "E18:F20", "E26:F29", "E35:F37", "E43:F45", "E51:F53", "E59:F59", "E65:F65", "F70:F72"];

    colA.getFont().setBold(true);
    rows.map((row, index) => {
        const currentRow = index + 1;

        if (bldCntr.includes(currentRow)) {
          const rng = sheet.getRange(`A${currentRow}:H${currentRow}`).getFormat();
          rng.getFont().setBold(true);
          rng.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        }
        if (yellow.includes(currentRow)) {
            const rng = sheet.getRange(`A${currentRow}`).getFormat();
            rng.getFill().setColor(Color.YELLOW);
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

        lblue.forEach(addy => {
          sheet.getRange(addy).getFormat().getFill().setColor(Color.LIGHTBLUE);
        });

        dblue.forEach(addy => {
          sheet.getRange(addy).getFormat().getFill().setColor(Color.DARKBLUE);
        });
    });

    grey.forEach(g => {
        sheet.getRange(`A${g}:F${g}`).getFormat().getFill().setColor(Color.LIGHTGREY);
    });
    sheet.getRange("B73:D73").getFormat().getFill().setColor(Color.DARKGREY);

    [sheet.getRange("A1:F12").getFormat(),
    sheet.getRange("A14:F21").getFormat(),
    sheet.getRange("A23:F29").getFormat(),
    sheet.getRange("A31:F37").getFormat(),
    sheet.getRange("A39:F45").getFormat(),
    sheet.getRange("A47:F53").getFormat(),
    sheet.getRange("A55:F59").getFormat(),
    sheet.getRange("A61:F65").getFormat(),
    sheet.getRange("A68:F73").getFormat(),
    sheet.getRange("A76:F76").getFormat(),
    sheet.getRange("A78:F79").getFormat(),
    sheet.getRange("A81:H83").getFormat(),
    sheet.getRange("A85:H87").getFormat(),
    sheet.getRange("A89:F111").getFormat()].forEach(borderRange => {
        borderRange.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
        borderRange.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin);
        borderRange.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin);
        borderRange.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin);
        borderRange.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin);
        borderRange.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(ExcelScript.BorderWeight.thin);
    });

    sheet.getUsedRange(true).getFormat().setWrapText(true);

    colA.setRowHeight(200);
    row1.setColumnWidth(80);
    sheet.getRange("G:G").getFormat().setColumnWidth(200);
    row1.autofitColumns();
    colA.autofitRows();
}