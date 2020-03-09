package org.apache.poi.excel.model;

import java.time.LocalDateTime;
import java.util.Date;

import org.apache.poi.excel.annotation.ExcelCell;
import org.apache.poi.excel.annotation.ExcelSheet;

@ExcelSheet(name = "Custom Sheet Name", heading = "Custom Sheet Heading")
public class ExcelAnnotated {

	@ExcelCell(header = "String Column")
	public String string;

	@ExcelCell(header = "Integer Column", type = ExcelCellType.INTEGER)
	public int integer;

	@ExcelCell(header = "Currency Column", type = ExcelCellType.CURRENCY)
	public int currency;

	@ExcelCell(header = "Decimal Column", type = ExcelCellType.DECIMAL)
	public float decimal;

	@ExcelCell(header = "Precise Column", type = ExcelCellType.PRECISE)
	public double precise;

	@ExcelCell(header = "Date Column", type = ExcelCellType.DATE)
	public Date date;

	@ExcelCell(type = ExcelCellType.DATETIME)
	public LocalDateTime dateTime;

	@ExcelCell(type = ExcelCellType.PERCENT)
	public float percent;

	/**
	 * @param string
	 * @param integer
	 * @param currency
	 * @param decimal
	 * @param precise
	 * @param date
	 * @param dateTime
	 * @param percent
	 */
	public ExcelAnnotated(String string, int integer, int currency, float decimal, double precise, Date date,
			LocalDateTime dateTime, float percent) {
		this.string = string;
		this.integer = integer;
		this.currency = currency;
		this.decimal = decimal;
		this.precise = precise;
		this.date = date;
		this.dateTime = dateTime;
		this.percent = percent;
	}
}
