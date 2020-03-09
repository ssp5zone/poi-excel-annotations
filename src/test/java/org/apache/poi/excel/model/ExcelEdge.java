package org.apache.poi.excel.model;

import org.apache.poi.excel.annotation.ExcelCell;
import org.apache.poi.excel.annotation.ExcelSheet;

@ExcelSheet(name = "Edge Cases", heading = "Test of Edge Cases")
public class ExcelEdge {

	@ExcelCell(header = "String Date", type = ExcelCellType.DATE)
	private String date;

	@ExcelCell(header = "String Date Time", type = ExcelCellType.DATETIME)
	private String dateTime;

	@ExcelCell(header = "String Currency", type = ExcelCellType.CURRENCY)
	private String currency;

	@ExcelCell(header = "String Percent", type = ExcelCellType.PRECISE)
	private String precise;

	public ExcelEdge(String date, String dateTime, String currency, String precise) {
		this.date = date;
		this.dateTime = dateTime;
		this.currency = currency;
		this.precise = precise;
	}
}
