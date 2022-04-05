function main(workbook: ExcelScript.Workbook) {
	let selectedSheet = workbook.addWorksheet();
	let date = new Date(Date.now());
	// Set range A1:B1 on selectedSheet
	selectedSheet.getRange("A1:H1").setValues([["DATE", "DAY OF WEEK", "S-T(DAY)", "S-T(NIGHT)", "W-F(DAY)", "W-F(NIGHT", "WEEKLY SHIFT TOTAL(DAY)", "WEEKLY SHIFT TOTAL(NIGHT)"]]);
	// Set width of column(s) at range B:B on selectedSheet to 67.5
	selectedSheet.getRange("B:B").getFormat().setColumnWidth(67.5);
	// Set width of column(s) at range D:D on selectedSheet to 57.75
	selectedSheet.getRange("D:D").getFormat().setColumnWidth(57.75);
	// Set width of column(s) at range F:F on selectedSheet to 64.5
	selectedSheet.getRange("F:F").getFormat().setColumnWidth(64.5);
	// Set width of column(s) at range G:G on selectedSheet to 129.75
	selectedSheet.getRange("G:G").getFormat().setColumnWidth(129.75);
	// Set width of column(s) at range H:H on selectedSheet to 129.75
	selectedSheet.getRange("H:H").getFormat().setColumnWidth(140);
	// Apply cell style on range A1:H1 on selectedSheet
	selectedSheet.getRange("A1:H1").setPredefinedCellStyle("Heading1");
	// Set font size to 11 for range A1:H1 on selectedSheet
	selectedSheet.getRange("A1:H1").getFormat().getFont().setSize(11);
	// Sets initial date, middle section needs var
	selectedSheet.getRange("B35").setValue(date.toLocaleDateString());
	// Sets  B39 to first day of month
	selectedSheet.getRange("B39").setValue("=EOMONTH(B35,-1)+1");
	// Sets firstOfMonth to B39's value
	let firstOfMonth = selectedSheet.getRange("B39").getValue();
	// Sets A2 to first day of month
	selectedSheet.getRange("A2").setValue(firstOfMonth);
	// Set number format for range A2 on selectedSheet
	selectedSheet.getRange("A2").setNumberFormatLocal("[$-en-US]dd-mmm-yy;@");
	// Sets A38 to days in current month
	selectedSheet.getRange("A38").setValue("=DAY(EOMONTH(A2,0))");
	// Sets y to the number of days to autofill
	let y = selectedSheet.getRange("A38").getValue();
	selectedSheet.getRange("B38").setValue("=" + y + "+1");
	let fillDateA = "A2:A" + selectedSheet.getRange("B38").getValue();
	let fillDateB = "B2:B" + selectedSheet.getRange("B38").getValue();
	// Auto fill range for dates
	selectedSheet.getRange("A2").autoFill(fillDateA, ExcelScript.AutoFillType.fillDefault);
	// Sets B2 cell to =TEXT(A2,"dddd")
	selectedSheet.getRange("B2").setFormulaLocal("=TEXT(A2,\"dddd\")");
	// Auto fill range for days of week
	selectedSheet.getRange("B2").autoFill(fillDateB, ExcelScript.AutoFillType.fillDefault);
	// this offsets the range selectedSheet.getRange("C2").getOffsetRange(1,0).setValue("hi");
	// Sets A35 to 1-Aug-2021
	selectedSheet.getRange("A35").setValue("01-Aug-2021");
	// Sets A36 to number of days since 1-Aug-21
	selectedSheet.getRange("A36").setValue("=A2-A35");
	// Set B36
	let h = selectedSheet.getRange("A36").getValue();
	selectedSheet.getRange("B36").setValue("=" + h);
	// Sets A37 to A36 mod 14
	selectedSheet.getRange("A37").setValue("=MOD(B36,14)");
	// Sets x to be counter
	let x = 0;
	// Sets start to start of week
	let start: number = 2;
	let end: number = 2;
	selectedSheet.getRange("B36").setValue("=" + h + "-1");
	h = selectedSheet.getRange("B36").getValue();
	for (x = 0; x < y; x++) {
		selectedSheet.getRange("B36").setValue("=" + h + "+1");
		h = selectedSheet.getRange("B36").getValue();
		switch (selectedSheet.getRange("A37").getValue()) {
			case 0:
				// Set fill color to D0CECE for range E2:F2 on selectedSheet
				selectedSheet.getRange("C2:D2").getOffsetRange(x, 2).getFormat().getFill().setColor("D0CECE");
				break;
			case 1:
				selectedSheet.getRange("C2:D2").getOffsetRange(x, 2).getFormat().getFill().setColor("D0CECE");
				break;
			case 2:
				selectedSheet.getRange("C2:D2").getOffsetRange(x, 2).getFormat().getFill().setColor("D0CECE");
				end = x + 2;
				selectedSheet.getRange("C2").getOffsetRange(x, 4).setValue("=SUM(C" + start + ":C" + end + ")");
				selectedSheet.getRange("C2").getOffsetRange(x, 5).setValue("=SUM(D" + start + ":D" + end + ")");
				start = end + 1;
				break;
			case 3:
				selectedSheet.getRange("C2:D2").getOffsetRange(x, 0).getFormat().getFill().setColor("D0CECE");
				break;
			case 4:
				selectedSheet.getRange("C2:D2").getOffsetRange(x, 0).getFormat().getFill().setColor("D0CECE");
				break;
			case 5:
				selectedSheet.getRange("C2:D2").getOffsetRange(x, 0).getFormat().getFill().setColor("D0CECE");
				end = x + 2;
				selectedSheet.getRange("C2").getOffsetRange(x, 4).setValue("=SUM(E" + start + ":E" + end + ")");
				selectedSheet.getRange("C2").getOffsetRange(x, 5).setValue("=SUM(F" + start + ":F" + end + ")");
				start = end + 1;
				break;
			case 6:
				selectedSheet.getRange("C2:D2").getOffsetRange(x, 2).getFormat().getFill().setColor("D0CECE");
				break;
			case 7:
				selectedSheet.getRange("C2:D2").getOffsetRange(x, 2).getFormat().getFill().setColor("D0CECE");
				break;
			case 8:
				selectedSheet.getRange("C2:D2").getOffsetRange(x, 2).getFormat().getFill().setColor("D0CECE");
				break;
			case 9:
				selectedSheet.getRange("C2:D2").getOffsetRange(x, 2).getFormat().getFill().setColor("D0CECE");
				end = x + 2;
				selectedSheet.getRange("C2").getOffsetRange(x, 4).setValue("=SUM(C" + start + ":C" + end + ")");
				selectedSheet.getRange("C2").getOffsetRange(x, 5).setValue("=SUM(D" + start + ":D" + end + ")");
				start = end + 1;
				break;
			case 10:
				selectedSheet.getRange("C2:D2").getOffsetRange(x, 0).getFormat().getFill().setColor("D0CECE");
				break;
			case 11:
				selectedSheet.getRange("C2:D2").getOffsetRange(x, 0).getFormat().getFill().setColor("D0CECE");
				break;
			case 12:
				selectedSheet.getRange("C2:D2").getOffsetRange(x, 0).getFormat().getFill().setColor("D0CECE");
				break;
			case 13:
				selectedSheet.getRange("C2:D2").getOffsetRange(x, 0).getFormat().getFill().setColor("D0CECE");
				end = x + 2;
				selectedSheet.getRange("C2").getOffsetRange(x, 4).setValue("=SUM(E" + start + ":E" + end + ")");
				selectedSheet.getRange("C2").getOffsetRange(x, 5).setValue("=SUM(F" + start + ":F" + end + ")");
				start = end + 1;
				break;
		}
	}
	if ((selectedSheet.getRange("A37").getValue() < 3) || (selectedSheet.getRange("A37").getValue() > 5 && selectedSheet.getRange("A37").getValue() < 10)) {
		selectedSheet.getRange("C2").getOffsetRange(x, 4).setValue("=SUM(C" + start + ":C" + end + ")");
		selectedSheet.getRange("C2").getOffsetRange(x, 5).setValue("=SUM(D" + start + ":D" + end + ")");
	}
	if ((selectedSheet.getRange("A37").getValue() > 2 && selectedSheet.getRange("A37").getValue() < 6) || (selectedSheet.getRange("A37").getValue() > 9)) {
		selectedSheet.getRange("C2").getOffsetRange(x, 4).setValue("=SUM(E" + start + ":E" + end + ")");
		selectedSheet.getRange("C2").getOffsetRange(x, 5).setValue("=SUM(F" + start + ":F" + end + ")");
	}
	// Set range J1:K4 on selectedSheet
	selectedSheet.getRange("J1:K4").setFormulasLocal([["W-F NIGHT", "=SUM(F2:F32)"], ["S-T NIGHT", "=SUM(D2:D32)"], ["W-F DAY", "=SUM(E2:E32)"], ["S-T DAY", "=SUM(C2:C32)"]]);
	// Set width of column(s) at range J:J on selectedSheet to 58.5
	selectedSheet.getRange("J:J").getFormat().setColumnWidth(100);
	selectedSheet.getRange("K:K").getFormat().setColumnWidth(100);
	// Add a new table at range J1:K4 on selectedSheet
	let newTable = workbook.addTable(selectedSheet.getRange("J1:K4"), false);
	// Set range J1:K1 on selectedSheet
	selectedSheet.getRange("J1:K1").setValues([["SHIFT", "TOTAL"]]);
	// Set range J7:K7 on selectedSheet
	selectedSheet.getRange("J7:K7").setFormulasLocal([["GRAND TOTAL", "=SUM(K2:K5)"]]);
	// Set font bold to true for range J7:K7 on selectedSheet
	selectedSheet.getRange("J7:K7").getFormat().getFont().setBold(true);
	// Apply cell style on range J7:K7 on selectedSheet
	selectedSheet.getRange("J7:K7").setPredefinedCellStyle("Total");
	// Insert chart on sheet selectedSheet
	let chart_1 = selectedSheet.addChart(ExcelScript.ChartType.barClustered, selectedSheet.getRange("J2:K5"));
	// Resize and move chart chart_1
	chart_1.setLeft(657.75);
	chart_1.setTop(134.25);
	chart_1.setWidth(360);
	chart_1.setHeight(216);
	// Sets A40 and B40 to MONTH and YEAR
	selectedSheet.getRange("A40").setValue("=UPPER(TEXT(A2,\"mmmmmmmmmm\"))");
	selectedSheet.getRange("B40").setValue("=TEXT(A2,\"yyyy\")");
	// Changes sheet name to current month and year
	let monthName = selectedSheet.getRange("A40").getValue();
	let yearNum = selectedSheet.getRange("B40").getValue();
	selectedSheet.setName(monthName + " " + yearNum);
	// Clear ExcelScript.ClearApplyTo.contents from range A35:B38 on selectedSheet
	selectedSheet.getRange("A35:B40").clear(ExcelScript.ClearApplyTo.contents);
	// Sets J33:M33 to total month values per shift
	selectedSheet.getRange("J33").setFormulaLocal("=SUM(C2:C32)");
	// Auto fill range
	selectedSheet.getRange("J33").autoFill("J33:M33", ExcelScript.AutoFillType.fillDefault);
	selectedSheet.getRange("J33:M33").getFormat().getFont().setColor("FFFFFF");
	// Changes sheet position to second to last
	let wsCount = workbook.getWorksheets().length;
	selectedSheet.setPosition(wsCount - 2);
}