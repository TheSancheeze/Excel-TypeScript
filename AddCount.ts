function main(workbook: ExcelScript.Workbook) {
	// Add a new worksheet
	let m = workbook.getLastWorksheet();
	let wsCount = workbook.getWorksheets().length;
	let x = 0;
	let j33 = "J33";
	let k33 = "K33";
	let l33 = "L33";
	let m33 = "M33";
	let n = "=SUM(";
	let o = "=SUM(";
	let l = "=SUM(";
	let k = "=SUM(";
	for (x = 0; x < (wsCount-1); x++) {
		let z = workbook.getWorksheets()[x];
		let p = z.getName();
		n = n + "'" + p + "'!" + j33 + ",";
		o = o + "'" + p + "'!" + k33 + ",";
		l = l + "'" + p + "'!" + l33 + ",";
		k = k + "'" + p + "'!" + m33 + ",";
	}
	n = n + ")";
	o = o + ")";
	l = l + ")";
	k = k + ")";
	m.getRange("B1").setValue(n);
	m.getRange("B2").setValue(o);
	m.getRange("B3").setValue(l);
	m.getRange("B4").setValue(k);
}