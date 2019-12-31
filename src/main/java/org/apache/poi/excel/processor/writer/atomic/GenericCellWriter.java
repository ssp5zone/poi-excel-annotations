package org.apache.poi.excel.processor.writer.atomic;

import java.lang.reflect.Field;
import java.util.Calendar;
import java.util.Date;
import java.util.function.BiConsumer;

import org.apache.poi.excel.model.ExcelCellType;
import org.apache.poi.excel.model.WorkbookContainer;
import org.apache.poi.excel.processor.reader.FieldReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class GenericCellWriter extends FieldReader {
	private final static Logger log = LoggerFactory.getLogger(GenericCellWriter.class);
	private WorkbookContainer container;

	public GenericCellWriter(Field field, WorkbookContainer container) {
		super(field);
		this.container = container;
	}

	public BiConsumer<Cell, Object> intWriter = (Cell cell, Object obj) -> {
		cell.setCellValue(this.getInt(obj));
		cell.setCellStyle(container.getStyle(ExcelCellType.INTEGER));
	};

	public BiConsumer<Cell, Object> shortWriter = (Cell cell, Object obj) -> {
		cell.setCellValue(this.getShort(obj));
		cell.setCellStyle(container.getStyle(ExcelCellType.INTEGER));
	};

	public BiConsumer<Cell, Object> longWriter = (Cell cell, Object obj) -> {
		cell.setCellValue(this.getLong(obj));
		cell.setCellStyle(container.getStyle(ExcelCellType.INTEGER));
	};

	public BiConsumer<Cell, Object> doubleWriter = (Cell cell, Object obj) -> {
		cell.setCellValue(this.getDouble(obj));
		cell.setCellStyle(container.getStyle(ExcelCellType.PRECISE));
	};

	public BiConsumer<Cell, Object> floatWriter = (Cell cell, Object obj) -> {
		cell.setCellValue(this.getFloat(obj));
		cell.setCellStyle(container.getStyle(ExcelCellType.DECIMAL));
	};

	public BiConsumer<Cell, Object> byteWriter = (Cell cell, Object obj) -> {
		cell.setCellValue(this.getByte(obj));
		cell.setCellStyle(container.getStyle(ExcelCellType.INTEGER));
	};

	@SuppressWarnings("deprecation")
	public BiConsumer<Cell, Object> charWriter = (Cell cell, Object obj) -> {
		cell.setCellValue(String.valueOf(this.getChar(obj)));
		cell.setCellType(CellType.STRING);
	};

	@SuppressWarnings("deprecation")
	public BiConsumer<Cell, Object> booleanWriter = (Cell cell, Object obj) -> {
		cell.setCellValue(this.getBoolean(obj));
		cell.setCellType(CellType.BOOLEAN);
	};

	@SuppressWarnings("deprecation")
	public BiConsumer<Cell, Object> utilDateWriter = (Cell cell, Object obj) -> {
		Date value = null;
		try {
			value = (Date) field.get(obj);

		} catch (IllegalArgumentException | IllegalAccessException | NullPointerException | ClassCastException e) {
			log.warn("Unable to write cell : " + cell + ". Defaulting to ERROR.");
			cell.setCellType(CellType.ERROR);
		}
		cell.setCellValue(value);
		cell.setCellStyle(container.getStyle(ExcelCellType.DATE));
	};

	@SuppressWarnings("deprecation")
	public BiConsumer<Cell, Object> sqlDateWriter = (Cell cell, Object obj) -> {
		Date value = null;
		try {
			value = new Date(((java.sql.Date) field.get(obj)).getTime());
			cell.setCellStyle(container.getStyle(ExcelCellType.DATE));
		} catch (IllegalArgumentException | IllegalAccessException | NullPointerException | ClassCastException e) {
			log.warn("Unable to write cell : " + cell + ". Defaulting to ERROR.");
			cell.setCellType(CellType.ERROR);
		}
		cell.setCellValue(value);
	};

	@SuppressWarnings("deprecation")
	public BiConsumer<Cell, Object> calendarWriter = (Cell cell, Object obj) -> {
		Date value = null;
		try {
			value = ((Calendar) field.get(obj)).getTime();
			cell.setCellStyle(container.getStyle(ExcelCellType.DATETIME));
		} catch (IllegalArgumentException | IllegalAccessException | NullPointerException | ClassCastException e) {
			log.warn("Unable to write cell : " + cell + ". Defaulting to ERROR.");
			cell.setCellType(CellType.ERROR);
		}
		cell.setCellValue(value);
	};

	@SuppressWarnings("deprecation")
	public BiConsumer<Cell, Object> stringWriter = (Cell cell, Object obj) -> {
		obj = this.getObject(obj);
		if (obj != null) {
			cell.setCellValue(obj.toString());
			cell.setCellType(CellType.STRING);
		}
	};
}
