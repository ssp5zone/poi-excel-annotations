package org.apache.poi.excel.processor.reader;

import java.lang.reflect.Field;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * A generalization of the Reflection - translation utilities.
 * 
 * @author ssp5zone
 */
public class FieldReader {
	private final static Logger log = LoggerFactory.getLogger(FieldReader.class);
	protected Field field;

	public FieldReader(Field field) {
		this.field = field;
	}

	/**
	 * Gets the generic object from a reflection field. <br>
	 * The only reason for writing this separately was to reduce the exception
	 * checks redundancies.
	 * 
	 * @param obj The object
	 * @return The object
	 */
	public Object getObject(Object obj) {
		try {
			return field.get(obj);
		} catch (IllegalArgumentException | IllegalAccessException | NullPointerException e) {
			log.warn("Unable to read field from the paseed excel object. The error was: ", e);
			return null;
		}
	}

	public int getInt(Object obj) {
		try {
			return field.getInt(obj);
		} catch (IllegalArgumentException | IllegalAccessException | NullPointerException e) {
			log.warn("Unable to read field from the paseed excel object. The error was: ", e);
			;
			return 0;
		}
	}

	public float getFloat(Object obj) {
		try {
			return field.getFloat(obj);
		} catch (IllegalArgumentException | IllegalAccessException | NullPointerException e) {
			log.warn("Unable to read field from the paseed excel object. The error was: ", e);
			;
			return 0;
		}
	}

	public double getDouble(Object obj) {
		try {
			return field.getDouble(obj);
		} catch (IllegalArgumentException | IllegalAccessException | NullPointerException e) {
			log.warn("Unable to read field from the paseed excel object. The error was: ", e);
			;
			return 0;
		}
	}

	public long getLong(Object obj) {
		try {
			return field.getLong(obj);
		} catch (IllegalArgumentException | IllegalAccessException | NullPointerException e) {
			log.warn("Unable to read field from the paseed excel object. The error was: ", e);
			;
			return 0;
		}
	}

	public short getShort(Object obj) {
		try {
			return field.getShort(obj);
		} catch (IllegalArgumentException | IllegalAccessException | NullPointerException e) {
			log.warn("Unable to read field from the paseed excel object. The error was: ", e);
			;
			return 0;
		}
	}

	public byte getByte(Object obj) {
		try {
			return field.getByte(obj);
		} catch (IllegalArgumentException | IllegalAccessException | NullPointerException e) {
			log.warn("Unable to read field from the paseed excel object. The error was: ", e);
			;
			return 0;
		}
	}

	public char getChar(Object obj) {
		try {
			return field.getChar(obj);
		} catch (IllegalArgumentException | IllegalAccessException | NullPointerException e) {
			log.warn("Unable to read field from the paseed excel object. The error was: ", e);
			;
			return ' ';
		}
	}

	public boolean getBoolean(Object obj) {
		try {
			return field.getBoolean(obj);
		} catch (IllegalArgumentException | IllegalAccessException | NullPointerException e) {
			log.warn("Unable to read field from the paseed excel object. The error was: ", e);
			;
			return false;
		}
	}
}
