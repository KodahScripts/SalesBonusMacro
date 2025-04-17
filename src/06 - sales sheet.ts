class SalesSheet {
  constructor(workbook: ExcelScript.Workbook, employee: Employee) {
    if (employee.deals.length == 0) return;
    const sheet = workbook.addWorksheet(employee.name);

    const data_lastRow = employee.getReportResultRow()
    const report_lastRow = employee.getReportResultRow() - 1;
    const results_row1 = employee.getResultRow(1);
    const results_row2 = employee.getResultRow(2);
    const results_row3 = employee.getResultRow(3);
    const results_row4 = employee.getResultRow(4);
    const results_row5 = employee.getResultRow(5);
    const results_row6 = employee.getResultRow(6);
    const results_row7 = employee.getResultRow(7);
    const results_row8 = employee.getResultRow(8);
    const results_row9 = employee.getResultRow(9);
    const results_row10 = employee.getResultRow(10);
    const results_row11 = employee.getResultRow(11);
    const results_row12 = employee.getResultRow(12);
    const results_row13 = employee.getResultRow(13);
    const results_row14 = employee.getResultRow(14);
    const results_row15 = employee.getResultRow(15);
    const results_row16 = employee.getResultRow(16);

    const headerRange: ExcelScript.Range = sheet.getRange("A1:B6");
    const colHeaderRange: ExcelScript.Range = sheet.getRange("A7:P7"); headerRange;
    const reportRange: ExcelScript.Range = sheet.getRange(`A8:R${data_lastRow}`);

    let reportData: Array<string | number>[] = [];

    employee.deals.forEach(deal => {
      reportData.push([deal.date, deal.id, deal.customer.id, deal.customer.name, deal.vehicle.id, deal.vehicle.year, deal.vehicle.make, deal.vehicle.model, deal.vehicle.salesType, deal.commission.gross, deal.commission.grossPercentage, deal.unitCount, deal.commission.amount]);
    });

    headerRange.setValues([
      ["Name", employee.name],
      ["Employee Number", employee.id],
      ["90 Day Rolling Average #", `=VLOOKUP(B2, '90'!A:E, 5, 0)`],
      ["CSI", `=IFERROR(VLOOKUP(B2, 'NPS'!B:H, 7, 0), 0)`],
      ["# of Surveys", `=IFERROR(VLOOKUP(B2, 'NPS'!B:H, 3, 0), 0)`],
      ["Retro Percentage", `=VLOOKUP(${employee.getTotalUnits()}, 'Look Up Table'!A:B, 2, TRUE)`]
    ]);
    headerRange.getFormat().getFont().setBold(true);
    sheet.getRange("B1:B6").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right)
    sheet.getCell(5, 1).setNumberFormat("0%")

    colHeaderRange.setValues([["Date", "Reference #", "Customer #", "Customer Name", "Stock #", "Year", "Make", "Model", "Sale Type", "Commission F&I", "Commission Gross", "Units", "Commission Amount", "Retro Mini", "Retro Owed", "Retro Commission Payout"]]);
    colHeaderRange.getFormat().getFill().setColor("lightgrey");
    colHeaderRange.getFormat().setRowHeight(50);
    colHeaderRange.getFormat().getFont().setBold(true);
    colHeaderRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    colHeaderRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);

    reportData.forEach((row, index) => {
      const lineNumber = index + 8;
      sheet.getRange(`A${lineNumber}:M${lineNumber}`).setValues([row]);
      sheet.getRange(`N${lineNumber}:P${lineNumber}`).setValues([[
        `=IF(M${lineNumber}<=251, VLOOKUP(B3, 'Look Up Table'!I:J, 2, TRUE) * L${lineNumber}, 0)`,
        `=IF(N${lineNumber}>0, N${lineNumber} - M${lineNumber}, 0)`,
        `=IF(N${lineNumber} = 0, K${lineNumber} * B6, 0)`
      ]])
    })

    sheet.getRange(`J${data_lastRow}:P${data_lastRow}`).setValues([[
      `=SUM(J8:J${report_lastRow})`,
      `=SUM(K8:K${report_lastRow})`,
      `=SUM(L8:L${report_lastRow})`,
      `=SUM(M8:M${report_lastRow})`,
      `=SUM(N8:N${report_lastRow})`,
      `=SUM(O8:O${report_lastRow})`,
      `=SUM(P8:P${report_lastRow})`,
    ]]);

    sheet.getRange(`J${results_row1}:M${results_row1}`).setValues([[
      "Prior Draw Balance", '', '', `=-VLOOKUP(B2, '3213'!A:G, 7, 0)`
    ]]);

    sheet.getRange(`J${results_row2}:M${results_row2}`).setValues([[
      "Commission", 0.18, '', `=M${data_lastRow}`
    ]]);
    sheet.getRange(`K${results_row2}`).setNumberFormat("0%")

    sheet.getRange(`J${results_row3}:M${results_row3}`).setValues([[
      "Retro Commission", '=B6', '', `=P${data_lastRow}`
    ]]);

    sheet.getRange(`J${results_row4}:M${results_row4}`).setValues([[
      "Retro MINI", '', '', `=O${data_lastRow}`
    ]]);

    sheet.getRange(`J${results_row5}:M${results_row5}`).setValues([[
      "Total Retro Commission", '', '', `=SUM(P${data_lastRow}, O${data_lastRow})`
    ]]);

    sheet.getRange(`J${results_row6}:M${results_row6}`).setValues([[
      "Total F&I", '', '', `=J${data_lastRow}`
    ]]);

    sheet.getRange(`J${results_row7}:M${results_row7}`).setValues([[
      "25% Reserve F&I", -0.25, '', `=K${results_row7} * M${results_row6}`
    ]]);
    sheet.getRange(`K${results_row7}`).setNumberFormat("0%");

    sheet.getRange(`J${results_row8}:M${results_row8}`).setValues([[
      "Total F&I Payable Gross", '', '', `=M${results_row6} + M${results_row7}`
    ]]);

    sheet.getRange(`J${results_row9}:M${results_row9}`).setValues([[
      "Total F&I Payout", 0.05, '', `=K${results_row9} * M${results_row8}`
    ]]);
    sheet.getRange(`K${results_row9}`).setNumberFormat("0%");

    sheet.getRange(`J${results_row10}:M${results_row10}`).setValues([[
      "Top Salesman Bonus", "=VLOOKUP(B2,'Pay Summary'!A:D,4,0)", '', `=IF(K${results_row10} = 1, 500, 0)`
    ]]);

    sheet.getRange(`J${results_row11}:M${results_row11}`).setValues([[
      "Unit Bonus", `=L${data_lastRow}`, '', `=VLOOKUP(K${results_row11}, 'Look Up Table'!E:F, 2, TRUE)`
    ]]);

    sheet.getRange(`J${results_row12}:O${results_row12}`).setValues([[
      "CSI", '=B4', '', `=IF(B5>=3,IF(B4="3P",L${data_lastRow}*50,IF(B4="A",0,IF(B4="B",L${data_lastRow}*-50))),0)`, "Rolling 90 Day", `='NPS Sheet'!X51`
    ]]);
    sheet.getRange(`K${results_row12}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);

    sheet.getRange(`J${results_row13}:M${results_row13}`).setValues([[
      "Total Bonus", '', '', `=SUM(M${results_row11}, M${results_row12}, M${results_row10})`
    ]]);

    sheet.getRange(`J${results_row14}:M${results_row14}`).setValues([[
      "Spiff", '', '', `=IFERROR(VLOOKUP(B2,'SPIFFS'!A:H,8,0),0)`
    ]]);

    sheet.getRange(`J${results_row15}:M${results_row15}`).setValues([[
      "Total Pay", '', '', `=SUM(M${results_row2}, M${results_row5}, M${results_row9}, M${results_row13}, M${results_row1}, M${results_row14})`
    ]]);

    sheet.getRange(`J${results_row16}:M${results_row16}`).setValues([[
      "Bucket Total YTD", '', '', `=IF(M${results_row15}<0, SUM(M${results_row1}, M${results_row13}, M${results_row9}, M${results_row5}), 0)`
    ]]);

    sheet.getRange(`J${results_row1}:J${results_row16}`).getFormat().getFont().setBold(true);
    sheet.getRange(`M${results_row1}:M${results_row16}`).setNumberFormat(NumberFormat.ACCOUNTING);
    sheet.getRange(`N${results_row12}:O${results_row12}`).getFormat().getFont().setBold(true);

    const employeeSignatureRow = data_lastRow + 6;
    const managerSignatureRow = employeeSignatureRow + 6;

    sheet.getRange(`A${employeeSignatureRow}:C${employeeSignatureRow}`).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
    sheet.getRange(`D${employeeSignatureRow}`).setValue("Employee");

    sheet.getRange(`A${managerSignatureRow}:C${managerSignatureRow}`).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
    sheet.getRange(`D${managerSignatureRow}`).setValue("Manager");

    sheet.getRange(`J8:K${data_lastRow}`).setNumberFormat(NumberFormat.ACCOUNTING);
    sheet.getRange(`M8:P${data_lastRow}`).setNumberFormat(NumberFormat.ACCOUNTING);
    sheet.getRange(`A${data_lastRow}:P${data_lastRow}`).getFormat().getFill().setColor(Color.LIGHT_GREY);

    sheet.getRange("A:A").setNumberFormat(NumberFormat.DATE);
    reportRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    colHeaderRange.getFormat().autofitColumns();
  }
}
