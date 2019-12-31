package org.apache.poi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.excel.ExcelWriter;
import org.apache.poi.excel.model.ExcelCellType;

/**
 * An annotation that can be added to a Class Attribute to provide a custom
 * column name, column position and column type when using the
 * {@link ExcelWriter} utility. <br>
 * <br>
 * If you are using this, make sure to add the class level annotation
 * {@link ExcelSheet}.
 * <ol>
 * <li>If this annotation is <b>not added to any</b> attribute of the class
 * then,
 * <ul>
 * <li>All the class attributes regardless of scope are used in the Excel
 * file.</li>
 * <li>The <b>index</b> <i>(Column position)</i> is same as the order in which
 * the properties are defined.</li>
 * <li>The <b>header</b> <i>(Column name)</i> is the Camel case to space parsed
 * property name.
 * <li>The cell type is the object type</li>
 * </ul>
 * <li>If this annotation is <b>added to even one</b> attribute of the class
 * then,
 * <ul>
 * <li>Only the attribute with this annotations are added in the Excel file.
 * Others are skipped.</li>
 * <li>The Column position is based on <b>index</b> passed, else the order in
 * which the properties are defined.</li>
 * <li>The Column name is based on the <b>header</b> parameter, else it is the
 * Camel case parsed property name.
 * <li>The cell type is based on the <b>type</b> property if passed else the
 * object's own type.</li>
 * </ul>
 * </ol>
 *
 * @author ssp5zone
 * @see ExcelCellType
 * @see ExcelWriter
 * @see ExcelSheet
 */
@Target({ ElementType.FIELD })
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCell {

	/**
	 * Use this to define the position of the column. <br>
	 * <br>
	 * If skipped, the default column ordering is same as the ordering of the class
	 * properties.
	 * 
	 * @return int
	 */
	public int index() default 0;

	/**
	 * The column name for this property. <br>
	 * <br>
	 * If skipped, the name is same as this property's name <i>(Camel Case to space
	 * separated string)</i>.
	 * 
	 * @return String
	 */
	public String header() default "";

	/**
	 * Format the cell to the mentioned type. <br>
	 * <br>
	 * If skipped, the parser tries to format using the attribute's own java type.
	 * 
	 * @return ExcelCellType
	 * @see ExcelCellType
	 */
	public ExcelCellType type() default ExcelCellType.DEFAULT;
}
