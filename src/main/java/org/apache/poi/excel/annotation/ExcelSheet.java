package org.apache.poi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.excel.ExcelWriter;

/**
 * The class level annotation that can be added to a POJO. Use this to define a
 * custom sheet name and a sheet heading when using the {@link ExcelWriter}
 * utility. <br>
 * <br>
 * This is mandatory if you want to use the custom Cell (property) level
 * annotation {@link ExcelCell}. Else, you can skip it. In the case of later,
 * the sheet name would be that of a class and no header would be added.
 *
 * @author parihsau
 * @see ExcelWriter
 * @see ExcelCell
 */
@Target({ ElementType.TYPE })
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelSheet {

	/**
	 * The name of this Sheet.
	 * 
	 * @return The name of the Sheet
	 */
	public String name() default "";

	/**
	 * A BIG heading that would appear in first 2 lines of the Excel Sheet. <br>
	 * <br>
	 * This would also add one line of timestamp indicating when the file was
	 * generated.
	 * 
	 * @return The header printed in the Sheet
	 */
	public String heading() default "";
}
