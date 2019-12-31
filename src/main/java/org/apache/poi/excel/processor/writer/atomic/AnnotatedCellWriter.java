package org.apache.poi.excel.processor.writer.atomic;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.time.ZonedDateTime;
import java.util.Date;
import java.util.function.Function;

import org.apache.commons.lang.ObjectUtils;
import org.apache.poi.excel.processor.reader.FieldReader;
import org.apache.poi.excel.utility.DateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class AnnotatedCellWriter extends FieldReader {
	private final static Logger log = LoggerFactory.getLogger(AnnotatedCellWriter.class);
	private Function<Object, Object> numericConverter;
	private Function<Object, Date> dateConverter;

	public AnnotatedCellWriter(Field field) {
		super(field);
	}

	public void initDateConverter() {
		if (field.getType() == Date.class) {
			dateConverter = (Object obj) -> (Date) this.getObject(obj);
		} else if (field.getType() == LocalDate.class) {
			dateConverter = (Object obj) -> DateUtil.asDate((LocalDate) this.getObject(obj));
		} else if (field.getType() == LocalDateTime.class) {
			dateConverter = (Object obj) -> DateUtil.asDate((LocalDateTime) this.getObject(obj));
		} else if (field.getType() == OffsetDateTime.class) {
			dateConverter = (Object obj) -> DateUtil.asDate((OffsetDateTime) this.getObject(obj));
		} else if (field.getType() == ZonedDateTime.class) {
			dateConverter = (Object obj) -> DateUtil.asDate((ZonedDateTime) this.getObject(obj));
		} else {
			dateConverter = (Object obj) -> DateUtil
					.parse(ObjectUtils.defaultIfNull(this.getObject(obj), "").toString());
		}
	}

	public void initNumericConverter() {
		if (field.getType() == int.class || field.getType() == Integer.class) {
			numericConverter = (Object obj) -> this.getInt(obj);
		} else if (field.getType() == float.class || field.getType() == Float.class) {
			numericConverter = (Object obj) -> this.getFloat(obj);
		} else if (field.getType() == double.class || field.getType() == Double.class) {
			numericConverter = (Object obj) -> this.getDouble(obj);
		} else if (field.getType() == long.class || field.getType() == Long.class) {
			numericConverter = (Object obj) -> this.getLong(obj);
		} else {
			numericConverter = (Object obj) -> this.getObject(obj);
		}
	}

	public void writeNumeric(Cell cell, Object obj) {
		Object genericObject = numericConverter.apply(obj);
		if (genericObject == null) {
			log.debug("An Excel of numeric cell family is null. Not writing anything. Cell: " + cell);
		}
		try {
			cell.setCellValue(Double.parseDouble(genericObject.toString()));
		} catch (Exception cce) {
			log.debug("An Excel cell is not recognized as Integer. Not writing anything in this cell: " + cell);
		}
	}

	public void writeDate(Cell cell, Object obj) {
		try {
			Date value = dateConverter.apply(obj);
			cell.setCellValue(value);
		} catch (IllegalArgumentException e) {
			log.warn("Unable to write Date to an Excel cell : " + cell + ". Defaulting to blank.");
		}
	}
}
