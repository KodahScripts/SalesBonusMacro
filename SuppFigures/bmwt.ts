const YEAR = "25";




// Change the YEAR variable and hit 'Run' button.  Book built.






























































function main(workbook: ExcelScript.Workbook) {
    const topRows = [
        ["!! Attention !!", '','','','',''],
        ["Supplemental figures reportd as larger than their related accounts on the financial",'','','','',''],
        ["BMW AND OTHER (EXCLUDING MINI) SUPPLEMENT",'','','','',''],
        ['', '', '', '', '', ''],
        ["(1) DOC/ADMINISTRATIVE FEES INCLUDED IN OTHER INCOME", '', '', '', '', ''],
        ["Any documentary, processing, administrative fee or any hard pack dollars from the sale of new and used vehicles, including new and used F & I and extended service plan income allocated to Account 805. If you produce a separate statement removing these fees and adding those dollars to gross, you do not need to fill out this section.", '', '', '', '', ''],
        ['', '', '', '', "MONTH", "YEAR TO DATE"],
        ["New BMW Vehicles", '', '', '', '', "=E8"],
        ["New Other Vehicles", '', '', '', '', ''],
        ["New F&I", '', '', '', '', ''],
        ["Used CPO Vehicles", '', '', '', '', "=E11"],
        ["Used Non-CPO Vehicles", '', '', '', '', ''],
        ["Used Other Vehicles", '', '', '', '', ''],
        ["Used F&I", '', '', '', '', ''],
        ["TOTAL (MUST BE LESS THAN OR EQUAL TO ACCOUNT 805 OF YOUR FINANCIAL STATEMENT)", '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ["(2) LEASE TERMINATION INFORMATION", '', '', '', '', ''],
        ["Office Manager and F&I Manager to set up a procedure for tracking this information.  New terminations = vehicles originally leased when new.  Used terminations = vehicles originally leased when used (includes program cars).", '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", ''],
        ['', '', "NEW", "USED", "NEW", "USED"],
        ["A. # Leases Renewed (We Leased the Customer Another Vehicle)", '', '', '', "=C22", "=D22"],
        ["B. # Lease Customers Who Returned the Vehicle & Walked Away", '', '', '', "=C23", "=D23"],
        ["C. # Customers Who Purchased Lease Return Unit or Another Vehicle from Dealership", '', '', '', "=C24", "=D24"],
        ["D. TOTAL LEASE TERMINATIONS (A+B+C)", '', "=C22+C23+C24", "=D22+D23+D24", "=E22+E23+E24", "=F22+F23+F24"],
        ['', '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ["(3) VEHICLES FINANCED", '', '', '', '', ''],
        ["Retail lease = any contract not titled to the retail customer.", '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", ''],
        ['', '', "NEW", "USED", "NEW", "USED"],
        ["# Vehicles Financed (Excluding Leases)", '', '', '', "=C32", "=D32"],
        ["# Retail Vehicles Leased", '', '', '', "=C33", "=D33"],
        ["Total Finance/Lease Reserve Gross Income (Excluding Insurance)", '', '', '', "=C34", "=D34"],
        ["All Chargebacks of Finance Income(Excluding Insurance Chargebacks)", '', '', '', "=C35", "=D35"],
        ['', '', '', '', '', ''],
        ["(5) GAP POLICIES", '', '', '', '', ''],
        ["Gross income is the dealership gross/commission from the GAP insurance contract.(Do not include physical damage or liability commissions.)", '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", ''],
        ['', '', "NEW", "USED", "NEW", "USED"],
        ["# GAP Policies Sold", '', '', '', "=C41", "=D41"],
        ["GAP Gross Income", '', '', '', "=C42", "=D42"],
        ["All Chargebacks of GAP Insurance Income", '', '', '', "=C43", "=D43"],
        ['', '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ["(8) OTHER CONTRACTED PROTECTION PRODUCTS (Tire & Rim, Key Replacement, Etch, Lease End Wear Care, etc.)", '', '', '', '', ''],
        ["Any gross profit on the following contract term products (Tire & Rim, Key Replacement, Etch, Lease End Wear Care, etc.) sold by the F&I department (and not included in Vehicle Gross).", '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", ''],
        ['', '', "NEW", "USED", "NEW", "USED"],
        [" GP Other Contracted Protection Products Sold by F&I", '', '', '', "=C50", "=D50"],
        ['', '', '', '', '', ''],
        ["(9) OTHER NON-CONTRACTED MERCHANDISE/AFTERMARKET (Accessories, Paint/Fabric/Rust, 3M, etc.)", '', '', '', '', ''],
        ["Any gross profit on accessories or other merchandise that are not contract term based (Paint/Fabric/Rust, 3M, etc.) sold by the F&I department (and not included in Vehicle Gross).", '', '', '', '', ''],
        ['', '', "MONTH", '', "YEAR TO DATE", ''],
        ['', '', "NEW", "USED", "NEW", "USED"],
        [" GP Other Non-Contracted Accessories/Other Merchandise (Aftermarket) Sold by F&I", '', '', '', "=C56", "=D56"],
        ['', '', '', '', '', ''],
        ["(10) SERVICE & BODY SHOP LABOR ANALYSIS", '', '', '', '', ''],
        ["We will check your Hours reported in each labor category by dividing the Sales $$ on the statement by the Hours Reported. If the resulting calculated Effective Labor Rate is too high (> $200) or too low (< $25), the Hours provided WILL NOT be used. We will also automatically update any obvious decimal related errors", '', '', '', '', ''],
        ["BMW - SERVICE LABOR ANALYSIS", "POSTED LABOR RATE", "EFFECTIVE LABOR RATE - MONTH", "EFFECTIVE LABOR RATE - YEAR TO DATE", "HOURS SOLD - MONTH", "HOURS SOLD - YEAR TO DATE"],
        ["Customer", '', '', '', '', "=E61"],
        ["Warranty Claim", '', '', '', '', "=E62"],
        ["Internal", '', '', '', '', "=E63"],
        ["TOTAL HOURS SOLD", '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ["(17) INVENTORY", '', '', '', "CARS", "TRUCKS"],
        ["# of New Vehicles 91+ Days in Inventory", '', '', '', '', ''],
        ["# of Used Vehicles 61+ Days in Inventory", '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ["(27) INTERNET SALES INFORMATION - MONTH & YTD", '', '', '', '', ''],
        ["MONTH", '', '', '', '', ''],
        ['', '', '', '', "NEW VEHICLES", "USED VEHICLES"],
        ["Your Website Leads", '', '', '', '', ''],
        ["Manufacturers Leads", '', '', '', '', ''],
        ["Lead Provider Leads", '', '', '', '', ''],
        ["TOTAL LEADS", '', '', '', "=E75+E76+E77", "=F75+F76+F77"],
        ['', '', '', '', '', ''],
        ["# Internet Sales", '', '', '', '', ''],
        ["Total Internet Front Gross", '', '', '', '', ''],
        ["Total Internet F&I Gross", '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ["(27) INTERNET SALES INFORMATION - MONTH & YTD", '', '', '', '', ''],
        ["YEAR TO DATE", '', '', '', '', ''],
        ['', '', '', '', "NEW VEHICLES", "USED VEHICLES"],
        ["Your Website Leads", '', '', '', "=E75", "=F75"],
        ["Manufacturers Leads", '', '', '', "=E76", "=F76"],
        ["Lead Provider Leads", '', '', '', "=E77", "=F77"],
        ["TOTAL LEADS", '', '', '', "=E88+E89+E90", "=F88+F89+F90"],
        ['', '', '', '', '', ''],
        ["# Internet Sales", '', '', '', "=E80", "=F80"],
        ["Total Internet Front Gross", '', '', '', "=E81", "=F81"],
        ["Total Internet F&I Gross", '', '', '', "=E82", "=F82"]
    ];

    const bottomTable = [
        ["PERSONNEL HEADCOUNT AS OF CURRENT MONTH END",'','','','','','','',''],
        ["USE BEST ESTIMATES TO PRORATE PERSONNEL WHOSE EFFORTS BENEFIT MORE THAN ONE DEPARTMENT.",'','','','','','','',''],
        ['',"NEW","F&I","PRE-OWNED","SERVICE","PARTS","BODY","ADMIN","TOTAL"],
        ["Owners", '', '', '', '', '', '', 1,"=SUM(B107:H107)"],
        ["Managers - BMW", '', '', '', '', '', '', '',"=SUM(B108:H108)"],
        ["Managers - MINI", '', '', '', '', '', '', '',"=SUM(B109:H109)"],
        ["Managers - Other",'','','','','','','',"=SUM(B110:H110)"],
        ["Salespeople - BMW",'','','','','','','',"=SUM(B111:H111)"],
        ["Salespeople - MINI",'','','','','','','',"=SUM(B112:H112)"],
        ["Salespeople - Other",'','','','','','','',"=SUM(B113:H113)"],
        ["Parts Counterpeople - BMW",'','','','','','','',"=SUM(B114:H114)"],
        ["Parts Counterpeople - MINI",'','','','','','','',"=SUM(B115:H115)"],
        ["Parts Counterpeople - Other",'','','','','','','',"=SUM(B116:H116)"],
        ["Technician - BMW",'','','','','','','',"=SUM(B117:H117)"],
        ["Technician - MINI",'','','','','','','',"=SUM(B118:H118)"],
        ["Technician - Other",'','','','','','','',"=SUM(B119:H119)"],
        ["Service Advisors - BMW",'','','','','','','',"=SUM(B120:H120)"],
        ["Service Advisors - MINI",'','','','','','','',"=SUM(B121:H121)"],
        ["Service Advisors - Other",'','','','','','','',"=SUM(B122:H122)"],
        ["Admin & Clerical - BMW",'','','','','','',1,"=SUM(B123:H123)"],
        ["Admin & Clerical - MINI",'','','','','','','',"=SUM(B124:H124)"],
        ["Admin & Clerical - Other",'','','','','','','',"=SUM(B125:H125)"],
        ["TOTAL - BMW", "=SUM(B107:B125)", "=SUM(C107:C125)", "=SUM(D107:D125)", "=SUM(E107:E125)", "=SUM(F107:F125)", "=SUM(G107:G125)", "=SUM(H107:H125)", "=SUM(B126:H126)"],
        ["TOTAL - MINI",'','','','','','','',"=SUM(B127:H127)"],
        ["TOTAL - Other",'','','','','','','',"=SUM(B128:H128)"],
        ["TOTAL - ALL", "=SUM(B126:B128)", "=SUM(C126:C128)", "=SUM(D126:D128)", "=SUM(E126:E128)", "=SUM(F126:F128)", "=SUM(G126:G128)", "=SUM(H126:H128)", "=SUM(I126:I128)"]
    ];

    const tabs = ["01.", "02.", "03.", "04.", "05.", "06.", "07.", "08.", "09.", "10.", "11.", "12."];
    const formulaRows = [7, 10, 21, 22, 23, 31, 32, 33, 34, 40, 41, 42, 49, 55, 60, 61, 62, 84, 85, 86, 87, 88, 89, 92, 93, 94];
    let lastSheetName = "";
    tabs.forEach(tab => {
        const sheetName = tab + YEAR;
        const sheet = workbook.addWorksheet(sheetName);
        const month = Number(tab.split('.')[0]);
        switch(month) {
            case 1:
                sheet.getRangeByIndexes(0, 0, topRows.length, topRows[0].length).setValues(topRows);
                break;
            default:
                let rows = topRows.map((row, index) => {
                    return row.map((r, i) => {
                        if(index > 85) {
                            if (r.startsWith("=") && formulaRows.includes(index)) {
                                const cell = r.slice(1);
                                const cellNumber = Number(cell.slice(1)) + 13;
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
        sheet.getRangeByIndexes(103,0,bottomTable.length,bottomTable[0].length).setValues(bottomTable);

        lastSheetName = sheetName;
        stylize(sheet);
    });
}


function stylize(sheet: ExcelScript.Worksheet) {
    const rows = sheet.getUsedRange().getValues();
    const row1 = sheet.getRange("1:1").getFormat();
    const colA = sheet.getRange("A:A").getFormat();

    const lightBlue = "#00B0F0";
    const lightGrey = "#BFBFBF";
    const darkGrey = "#808080";

    const full = [1, 2, 3, 4, 5, 6, 16, 17, 18, 19, 26, 27, 28, 29, 36, 37, 38, 44, 45, 46, 47, 51, 52, 53, 57, 58, 59, 65, 66, 70, 71, 72, 73, 79, 83, 84, 85, 86, 92];
    const mid = [7, 8, 9, 10, 11, 12, 13, 14, 15, 67, 68, 69, 74, 75, 76, 77, 78, 80, 81, 82, 87, 88, 89, 90, 91, 93, 94, 95];
    const small = [20, 21, 22, 23, 24, 25, 30, 31, 32, 33, 34, 35, 20, 21, 22, 23, 24, 25, 30, 31, 32, 33, 34, 35, 39, 40, 41, 42, 43, 48, 49, 50, 54, 55, 56];
    const doubleMerge = [20, 30, 39, 48, 54];
    const blue = [1, 2, 3, 4, 5, 6, 18, 19, 28, 29, 37, 38, 46, 47, 52, 53, 58, 59, 60, 67, 72, 73, 85, 86];
    const bldCntr = [7, 21, 31, 40, 49, 55, 60, 74, 87];
    const bld = ['E67', 'F67'];
    const grey = ['E8', 'F8', 'E11', 'F11', 'C22:F25', 'C32:F35', 'C41:F43', 'C50:F50', 'C56:F56', 'B61:F63', 'A64:F64', 'E68:F69', 'E75:F77', 'E80:F82', 'E88:F90', 'E93:F95'];

    colA.getFont().setBold(true);
    rows.map((row, index) => {
        const currentRow = index + 1;

        if (blue.includes(currentRow)) {
            const rng = sheet.getRange(`A${currentRow}`).getFormat();
            rng.getFill().setColor(lightBlue);
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
    
    bld.forEach(b => {
        sheet.getRange(b).getFormat().getFont().setBold(true);
        sheet.getRange(b).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    });

    grey.forEach(g => {
        sheet.getRange(g).getFormat().getFill().setColor(lightGrey);
    });


    const tableTop = sheet.getRange("A104:I106").getFormat();
    sheet.getRange("A104:I105").merge(true);
    
    tableTop.getFill().setColor(lightBlue);
    sheet.getRange("A106:I106").getFormat().getFont().setBold(true);
    sheet.getRange("B64:D64").getFormat().getFill().setColor(darkGrey);
    sheet.getRange("B107:H125").getFormat().getFill().setColor(darkGrey);
    sheet.getRange("H107:H108").getFormat().getFill().clear();
    sheet.getRange("B108:F108").getFormat().getFill().setColor("yellow");
    sheet.getRange("B111:D111").getFormat().getFill().setColor("yellow");
    sheet.getRange("F114").getFormat().getFill().setColor("yellow");
    sheet.getRange("E117").getFormat().getFill().setColor("yellow");
    sheet.getRange("E120:F121").getFormat().getFill().setColor("yellow");
    sheet.getRange("E122").getFormat().getFill().clear();
    sheet.getRange("B123:F123").getFormat().getFill().setColor("yellow");
    sheet.getRange("H123").getFormat().getFill().clear();
    sheet.getRange("I106:I129").getFormat().getFill().setColor(lightGrey);
    sheet.getRange("A126:H129").getFormat().getFill().setColor(lightGrey);


    [sheet.getRange("A1:F15").getFormat(),
    sheet.getRange("A18:F25").getFormat(),
    sheet.getRange("A28:F35").getFormat(),
    sheet.getRange("A37:F43").getFormat(),
    sheet.getRange("A46:F50").getFormat(),
    sheet.getRange("A52:F56").getFormat(),
    sheet.getRange("A58:F64").getFormat(),
    sheet.getRange("A67:F69").getFormat(),
    sheet.getRange("A72:F82").getFormat(),
    sheet.getRange("A85:F95").getFormat(),
    sheet.getRange("A104:I129").getFormat()].forEach(borderRange => {
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