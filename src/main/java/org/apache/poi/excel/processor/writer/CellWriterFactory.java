package org.apache.poi.excel.processor.writer;

import java.lang.reflect.Field;
import java.util.Calendar;
import java.util.Date;
import java.util.function.BiConsumer;

import org.apache.poi.excel.ExcelWriter;
import org.apache.poi.excel.annotation.ExcelCell;
import org.apache.poi.excel.model.ExcelCellType;
import org.apache.poi.excel.model.WorkbookContainer;
import org.apache.poi.excel.processor.writer.atomic.AnnotatedCellWriter;
import org.apache.poi.excel.processor.writer.atomic.GenericCellWriter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Creates generators that are used to write individual cells of an Excel file.
 * These generators significantly improve the performance by reducing the
 * generic checks over and again for a specific column. <br>
 * <br>
 * Local class, to be used internally.
 *
 * @author ssp5zone
 * @see ExcelWriter
 */
public class CellWriterFactory {
	private final static Logger log = LoggerFactory.getLogger(CellWriterFactory.class);
	private WorkbookContainer container;

	public CellWriterFactory(WorkbookContainer container) {
		this.container = container;
	}

	/**
	 * Generates a small cell writer block that that can be efficiently used for
	 * multiple iterations with a very low footprint. NOTE: Here the decision is
	 * made based on the Data type of field
	 * 
	 * @param field The field based on which the writer is decided
	 * @return A lambda function that can be utilized by the caller
	 */
	public BiConsumer<Cell, Object> getGenericFieldWriter(Field field) {
		Class<?> fieldClass = field.getType();
		GenericCellWriter cellWriter = new GenericCellWriter(field, container);
		if (fieldClass == Integer.class || fieldClass == int.class) {
			return cellWriter.intWriter;
		} else if (fieldClass == Short.class || fieldClass == short.class) {
			return cellWriter.shortWriter;
		} else if (fieldClass == Long.class || fieldClass == long.class) {
			return cellWriter.longWriter;
		} else if (fieldClass == Double.class || fieldClass == double.class) {
			return cellWriter.doubleWriter;
		} else if (fieldClass == Float.class || fieldClass == float.class) {
			return cellWriter.floatWriter;
		} else if (fieldClass == Byte.class || fieldClass == byte.class) {
			return cellWriter.byteWriter;
		} else if (fieldClass == Character.class || fieldClass == char.class) {
			return cellWriter.charWriter;
		} else if (fieldClass == Boolean.class || fieldClass == boolean.class) {
			return cellWriter.booleanWriter;
		} else if (fieldClass == Date.class) {
			return cellWriter.utilDateWriter;
		} else if (fieldClass == Calendar.class) {
			return cellWriter.calendarWriter;
		} else if (fieldClass == java.sql.Date.class) {
			return cellWriter.sqlDateWriter;
		} else {
			return cellWriter.stringWriter;
		}
	}

	/**
	 * Generates a small cell writer block that that can be efficiently used for
	 * multiple iterations with a very low footprint. NOTE: Here the decision is
	 * made based on the ExcelCell Type Annotation of the field
	 * 
	 * @param field The field based on which the writer is decided
	 * @return A lambda function that can be utilized by the caller
	 */
	public BiConsumer<Cell, Object> getAnnotatedFieldWriter(Field field) {
		if (field.isAnnotationPresent(ExcelCell.class)) {
			ExcelCellType type = field.getAnnotation(ExcelCell.class).type();
			if (!type.equals(ExcelCellType.DEFAULT)) {
				AnnotatedCellWriter cellWriter = new AnnotatedCellWriter(field);
				switch (type) {
				case GENERAL:
					return ((Cell cell, Object obj) -> {
						String value = "";
						try {
							Object attrValue = field.get(obj);
							if (attrValue != null) {
								value = attrValue.toString();
							}
						} catch (IllegalArgumentException | IllegalAccessException e) {
							log.warn("Unable to write generic excel cell : " + cell + ". Defaulting to blank.");
						}
						cell.setCellValue(value);
						cell.setCellStyle(container.getStyle(ExcelCellType.GENERAL));
					});
				case INTEGER:
					cellWriter.initNumericConverter();
					return ((Cell cell, Object obj) -> {
						cellWriter.writeNumeric(cell, obj);
						cell.setCellStyle(container.getStyle(ExcelCellType.INTEGER));
					});
				case DECIMAL:
					cellWriter.initNumericConverter();
					return ((Cell cell, Object obj) -> {
						cellWriter.writeNumeric(cell, obj);
						cell.setCellStyle(container.getStyle(ExcelCellType.DECIMAL));
					});
				case PRECISE:
					cellWriter.initNumericConverter();
					return ((Cell cell, Object obj) -> {
						cellWriter.writeNumeric(cell, obj);
						cell.setCellStyle(container.getStyle(ExcelCellType.PRECISE));
					});
				case CURRENCY:
					cellWriter.initNumericConverter();
					return ((Cell cell, Object obj) -> {
						cellWriter.writeNumeric(cell, obj);
						cell.setCellStyle(container.getStyle(ExcelCellType.CURRENCY));
					});
				case DATE:
					cellWriter.initDateConverter();
					return ((Cell cell, Object obj) -> {
						cellWriter.writeDate(cell, obj);
						cell.setCellStyle(container.getStyle(ExcelCellType.DATE));
					});
				case DATETIME:
					cellWriter.initDateConverter();
					return ((Cell cell, Object obj) -> {
						cellWriter.writeDate(cell, obj);
						cell.setCellStyle(container.getStyle(ExcelCellType.DATETIME));
					});
				case PERCENT:
					cellWriter.initNumericConverter();
					return ((Cell cell, Object obj) -> {
						cellWriter.writeNumeric(cell, obj);
						cell.setCellStyle(container.getStyle(ExcelCellType.PERCENT));
					});
				default:
					break;
				}
			}
		}
		// If the annotation or type is missing
		return this.getGenericFieldWriter(field);
	}

	/**
	 * A simple column name writer with custom style
	 * 
	 * @return A lambda function that can be utilized by the caller to write column
	 *         headings
	 */
	public BiConsumer<Cell, String> getColumnWriter() {
		Workbook wb = this.container.getWorkbook();

		Font font = wb.createFont();
		font.setBold(true);
		font.setColor(IndexedColors.DARK_BLUE.getIndex());

		CellStyle style = wb.createCellStyle();
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBottomBorderColor(IndexedColors.BLUE1.getIndex());
		style.setFont(font);

		return (Cell cell, String header) -> {
			cell.setCellValue(header);
			cell.setCellStyle(style);
		};
	}

}
