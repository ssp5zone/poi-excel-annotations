package org.apache.poi.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.stream.Collectors;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.excel.annotation.ExcelCell;
import org.apache.poi.excel.annotation.ExcelSheet;
import org.apache.poi.excel.model.SheetContainer;
import org.apache.poi.excel.model.WorkbookContainer;
import org.apache.poi.excel.utility.FileUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import freemarker.template.utility.StringUtil;

/**
 * An Excel Utility that automatically converts any {@link List} of POJO objects
 * to an Excel file. <br>
 * <br>
 * Works better if you have added the {@link ExcelCell} and {@link ExcelSheet}
 * annotations in your POJO class. <br>
 * <br>
 * <b>Usage:</b>
 * <ol>
 * <li>Simple excel file with <b>one sheet</b>.
 * 
 * <pre>
 * {@code
 * List<Employee> myList;
 * ...
 * // some code to fill this list
 * File file = ExcelWriter.write(myList);
 * }
 * </pre>
 * 
 * </li>
 * <li>An excel file with <b>multiple sheets</b>.
 * 
 * <pre>
 * {@code
 * List<Employee> empList;
 * List<Car> carList;
 * List<Location> locationList;
 * List<Department> departmentList;
 * List<System> systemList;
 * ...
 * // some code to populate the lists
 * File file = ExcelWriter.write(empList, carList, locationList, departmentList, systemList);
 * // Each of list passed gets added at a separate field each.
 * }
 * </pre>
 * 
 * </li>
 * <li>Excel file with a <b>custom file name</b>.
 * 
 * <pre>
 * {@code
 * List<DCHero> dcHeros;
 * List<MarvelHero> marvelHeros;
 * ...
 * // some code to fill these lists
 * File file = ExcelWriter.write(filename, dcHeros, marvelHeros);
 * }
 * </pre>
 * 
 * </li>
 * <li>Excel file with at a <b>custom location</b>.
 * 
 * <pre>
 * {@code
 * List<DCHero> dcHeros;
 * List<MarvelHero> marvelHeros;
 * ...
 * // some code to fill these lists
 * File file = ExcelWriter.write(filename, pathToDir, dcHeros, marvelHeros);
 * }
 * </pre>
 * 
 * </li>
 * </ol>
 * <h1>NOTE:</h1>This utility automatically creates a backup file which has a
 * retention period of 60 days. <br>
 * You can specify a custom excel storage directory by either,
 * <ol>
 * <li>Setting globally once via a static call :
 * {@link ExcelWriter#setStorage(String)}</li>
 * <li>Passing a custom path during each function call :
 * {@link ExcelWriter#write(String, String, List...)}</li>
 * </ol>
 * For more details of how the data is formatted, refer {@link ExcelSheet}.
 * 
 * @author ssp5zone
 * @see ExcelCell
 * @see ExcelSheet
 */
public class ExcelWriter {
	private final static Logger log = LoggerFactory.getLogger(ExcelWriter.class);

	private static String BACKUP_STORAGE_PATH;
	private static final String TEMP_PATH = "/var/tmp/";
	private static final int BACKUP_FILE_RETENTION_DAYS = 60;

	private static WorkbookContainer workbookContainer;

	/**
	 * Initialize the file storage directory
	 */
	static {
		BACKUP_STORAGE_PATH = TEMP_PATH;
		workbookContainer = new WorkbookContainer();
	}

	/**
	 * Use this to set the default path where the excel file would be written as a
	 * backup/storage. <br>
	 * <br>
	 * Alternately, you can also set it by adding
	 * <b>UNSEC.UTILS.EXCEL.BACKUP.DIR</b> property in your application.
	 * 
	 * @param path The physical location where the excel files would be stored.
	 */
	public static void setStorage(String path) {
		BACKUP_STORAGE_PATH = (path == null || path.equals("")) ? TEMP_PATH : path.trim();
	}

	/**
	 * Creates an Excel Workbook based on the data. Each list of data passed is
	 * converted to it's own sheet. NOTE: It is advised to pass a filename, else a
	 * dummy name would be generated.
	 * 
	 * @param data A list of Plain old java objects. Each list passed gets converted
	 *             to its own sheet.
	 * @param <T>  The datatype contained by the list.
	 * @return The generated Excel file.
	 */
	@SafeVarargs
	public static <T> File write(List<? extends T>... data) {
		String dummyFileName = String.valueOf(System.currentTimeMillis());
		return ExcelWriter.write(dummyFileName, data);
	}

	/**
	 * Creates an Excel Workbook based on the data. Each list of data passed is
	 * converted to it's own sheet. The generated data is stored as the file name
	 * provided at a predefined path. The file retention period is 20 days.
	 * 
	 * @param fileName The name of the generated file.
	 * @param data     A list of Plain old java objects. Each list passed gets
	 *                 converted to its own sheet.
	 * @param <T>      The datatype contained by the list.
	 * @return The generated Excel file.
	 */
	@SafeVarargs
	public static <T> File write(String fileName, List<? extends T>... data) {
		return ExcelWriter.write(fileName, BACKUP_STORAGE_PATH, data);
	}

	/**
	 * Creates an Excel Workbook based on the data. Each list of data passed is
	 * converted to it's own sheet. The generated data is stored as the file name
	 * provided at a predefined path. The file retention period is 20 days.
	 * 
	 * @param fileName The name of the generated file.
	 * @param path     The path where the file is to be stored.
	 * @param data     A list of Plain old java objects. Each list passed gets
	 *                 converted to its own sheet.
	 * @param <T>      The datatype contained by the list.
	 * @return The generated Excel file.
	 */
	@SafeVarargs
	public synchronized static <T> File write(String fileName, String path, List<? extends T>... data) {
		List<List<?>> filteredData = Arrays.asList(data).stream().filter(nonEmptyData).collect(Collectors.toList());
		// If there is no data in any sheet, do not process further
		if (filteredData.size() > 0) {
			// This is important to ensure that we multiple Threads do not call this
			// at-once.
			// The context switching happening here is heavy and may cause the whole system
			// to lag.
			synchronized (workbookContainer) {
				// Reset the workbook
				ExcelWriter.initWorkBook();

				// Process each sheet one by one
				filteredData.forEach(list -> {
					createSheet.andThen(generateName).andThen(giveHeading).andThen(addColumns).andThen(writeData)
							.andThen(autoSizeColumns).andThen(freezePane).andThen(attachFilters).apply(list);
				});

				// Write to actual location
				return writeToFile(path, fileName);
			}
		}
		return null;
	}

	/**
	 * Initializes a new WorkBook.
	 */
	private static void initWorkBook() {
		ExcelWriter.workbookContainer = new WorkbookContainer();
	}

	/**
	 * A simple predicate to check for a non-empty list object
	 */
	private static Predicate<List<?>> nonEmptyData = data -> data != null && data.size() > 0;

	/**
	 * A function that creates a new sheet from the existing workbook.
	 */
	private static Function<List<?>, SheetContainer> createSheet = (List<?> data) -> {
		SheetContainer sheetContainer = new SheetContainer();
		sheetContainer.setSheet(workbookContainer.getWorkbook().createSheet());
		sheetContainer.setData(data);
		return sheetContainer;
	};

	/**
	 * Generate a Sheet Name based on the Excel annotations -> ExcelSheet.sheetName
	 * If no annotation or a name is found, just use the Class Name as is.
	 */
	private static Function<SheetContainer, SheetContainer> generateName = (SheetContainer sheetContainer) -> {
		Workbook workbook = workbookContainer.getWorkbook();
		Sheet sheet = sheetContainer.getSheet();

		String sheetName = "";

		try {
			// We have already filtered out stuff where that data size !> 0. So NPE wont
			// occur here.
			Class<?> _class = sheetContainer.getData().get(0).getClass();

			// See if the good people added an Excel Sheet annotation
			if (_class.isAnnotationPresent(ExcelSheet.class)) {
				// And by any chance gave it a name
				sheetName = _class.getAnnotation(ExcelSheet.class).name();

				// Also get the heading for future use
				sheetContainer.setHeading(_class.getAnnotation(ExcelSheet.class).heading());
			}
			// If there are no annotations or no one bothered to give a sheet name, just use
			// the Class Name
			if (sheetName.equals("")) {
				sheetName = parseCamelCase(_class.getSimpleName());
			}
			// If some genius has used an Anonymous class, then just use the index.
			if (sheetName.equals("")) {
				sheetName = "Sheet - ".concat(String.valueOf((workbook.getSheetIndex(sheet))));
			}

			workbook.setSheetName(workbook.getSheetIndex(sheet), sheetName);

		} catch (Exception e2) {
			log.error("Was unable to give sheet its name", e2);
		}

		return sheetContainer;
	};

	/**
	 * Add a simple 2 line description of Line 1 : What this sheet is? (Sheet's
	 * Name) Line 2 : When was this generated? (Current Time)
	 */
	private static Function<SheetContainer, SheetContainer> giveHeading = (SheetContainer sheetContainer) -> {
		Sheet sheet = sheetContainer.getSheet();
		String heading = sheetContainer.getHeading();

		if (!heading.equals("")) {
			try {
				// Line 1
				sheet.createRow(0).createCell(0).setCellValue(heading);

				// Add some styling to the header
				Workbook wb = workbookContainer.getWorkbook();

				Font font = wb.createFont();
				font.setBold(true);
				font.setColor(IndexedColors.DARK_BLUE.getIndex());
				font.setFontHeightInPoints((short) (heading.length() < 16 ? 18 : 16));

				CellStyle style = wb.createCellStyle();
				style.setFont(font);

				sheet.getRow(0).getCell(0).setCellStyle(style);

				// Line 2
				sheet.createRow(1).createCell(0).setCellValue("Generated on: " + Calendar.getInstance().getTime());

				// Merge cells to make them look decent.
				sheet.addMergedRegion(new CellRangeAddress(0, // first row (0-based)
						0, // last row (0-based)
						0, // first column (0-based)
						5 // last column (0-based)
				));
				sheet.addMergedRegion(new CellRangeAddress(1, // first row (0-based)
						1, // last row (0-based)
						0, // first column (0-based)
						5 // last column (0-based)
				));

				// A spacer row
				sheet.createRow(2).createCell(0);

			} catch (Exception e3) {
				log.error("Was unable to create sheet's header", e3);
			}
		}

		return sheetContainer;
	};

	/**
	 * Add Header Columns by either using the ExcelCell Annotations or just by
	 * parsing the pojo fields. !!IMPORTANT!! If even 1 @ExcelCell annotation is
	 * found, it would then keep only those fields that are annotated
	 */
	private static Function<SheetContainer, SheetContainer> addColumns = (SheetContainer sheetContainer) -> {
		Sheet sheet = sheetContainer.getSheet();
		List<?> data = sheetContainer.getData();

		int rowIndex = sheetContainer.getHeading().equals("") ? 0 : 3;

		// Create a new row after the Heading (give +1 blank space)
		Row row = sheet.createRow(rowIndex);

		try {
			// Get the POJO class of the listed data
			Class<?> _class = data.get(0).getClass();

			// Get all fields (public, protected, anything)
			Field fields[] = _class.getDeclaredFields();

			// Filter out those fields that have Excel Cell annotations
			List<Field> annotatedFields = Arrays.asList(fields).stream().filter((Field field) -> {
				field.setAccessible(true);
				return field.isAnnotationPresent(ExcelCell.class);
			}).collect(Collectors.toList());

			// Get a function that writes the columns
			BiConsumer<Cell, String> columnWriter = workbookContainer.getWriterFactory().getColumnWriter();

			// 0 based column index
			AtomicInteger columnIndex = new AtomicInteger();

			// A custom stub that can be recursively executed
			Consumer<String> addColumn = (String header) -> {
				// Add the header column
				Cell cell = row.createCell(columnIndex.get());
				columnWriter.accept(cell, header);

				// Set min width to make the column accessible
				sheet.setColumnWidth(columnIndex.getAndIncrement(), ((header.length() + 3) * 256) + 200);
			};

			// If at-least 1 annotation is present
			if (annotatedFields.size() > 0) {
				// Process only the fields that have annotations
				annotatedFields.forEach((Field field) -> {

					// Check if the header name is present in the annotation
					String header = field.getAnnotation(ExcelCell.class).header();

					// If not, use the object name itself
					if (header.equals("")) {
						header = parseCamelCase(field.getName());
					}

					// Add the column
					addColumn.accept(header);
				});
			} else {
				// Process Everything
				/**
				 * The following code looks redundant and is somewhat same as above But creating
				 * a common code would increase a couple of "if" checks which is degrading if
				 * the data size is more.
				 */
				Arrays.asList(fields).forEach(field -> {
					// As annotation is not present, use the object name itself
					String header = parseCamelCase(field.getName());

					// Add the column
					addColumn.accept(header);
				});
			}

		} catch (Exception e4) {
			log.error("Was Unable to add columns to sheet: " + sheet.getSheetName(), e4);
		}
		return sheetContainer;
	};

	/**
	 * The one responsible for writing actual data each cell.
	 */
	private static Function<SheetContainer, SheetContainer> writeData = (SheetContainer sheetContainer) -> {
		Sheet sheet = sheetContainer.getSheet();
		List<?> dataList = sheetContainer.getData();
		try {
			// Get the POJO class of the listed data
			Class<?> _class = dataList.get(0).getClass();

			// Get all fields (public, protected, anything)
			Field fields[] = _class.getDeclaredFields();

			// Filter out those fields that have Excel Cell annotations
			List<Field> fieldList = Arrays.asList(fields).stream().filter((Field field) -> {
				field.setAccessible(true);
				return field.isAnnotationPresent(ExcelCell.class);
			}).collect(Collectors.toList());

			// If no annotations are present
			// then use all the fields
			if (fieldList.size() == 0) {
				fieldList = Arrays.asList(fields);
			}

			// Create a local map of how each field should be effectively written.
			Map<Field, BiConsumer<Cell, Object>> fieldWriter = new HashMap<Field, BiConsumer<Cell, Object>>();
			fieldList.forEach(field -> {
				// Create a dynamic function that knows how to write this specific "type"
				// of column in the Excel based on the annotations or its data type.
				BiConsumer<Cell, Object> cellWriter = workbookContainer.getWriterFactory()
						.getAnnotatedFieldWriter(field);
				fieldWriter.put(field, cellWriter);
			});

			// Shift rows down to accommodate for the heading and the column headers
			int shiftIndex = sheetContainer.getHeading().equals("") ? 1 : 4;

			// Write data to each cell.
			for (int rowNum = 0; rowNum < dataList.size(); rowNum++) {
				// + 3 as Row0 and Row1 are filled with the heading. Row2 is a spacer.
				Row row = sheet.createRow(rowNum + shiftIndex);

				// Get whatever the field holds from the object
				Object data = dataList.get(rowNum);

				for (int colNum = 0; colNum < fieldList.size(); colNum++) {
					// Get the current field
					Field field = fieldList.get(colNum);

					try {
						// Create a new cell
						Cell cell = row.createCell(colNum);

						// write the data
						fieldWriter.get(field).accept(cell, data);

					} catch (Exception ex) {
						log.warn("Unable to write data to row: " + (rowNum + 1) + " cell: " + (colNum + 1)
								+ " of sheet: " + sheet.getSheetName(), ex);
					}
				}
			}

		} catch (Exception e) {
			log.error("Was Unable to write data to sheet: " + sheet.getSheetName(), e);
		}
		return sheetContainer;
	};

	/**
	 * Resizes all the columns to ensure that all the data becomes visible.
	 * !!DANGER!! : Very slow. Avoid using this.
	 */
	private static Function<SheetContainer, SheetContainer> autoSizeColumns = (SheetContainer sheetContainer) -> {
		Sheet sheet = sheetContainer.getSheet();

		// In case of SXSSFSheet, the row tracking is limited and hence cannot be used
		// to auto-size
		if (!(sheet instanceof SXSSFSheet)) {
			for (int column = 0; column < sheetContainer.getData().size(); column++) {
				sheet.autoSizeColumn(column, false);
			}
		}

		return sheetContainer;
	};

	/**
	 * Freeze first 4 rows if a heading is present, else only 1 row
	 */
	private static Function<SheetContainer, SheetContainer> freezePane = (SheetContainer sheetContainer) -> {
		Sheet sheet = sheetContainer.getSheet();
		int frozenRows = sheetContainer.getHeading().equals("") ? 1 : 4;
		sheet.createFreezePane(0, frozenRows);
		return sheetContainer;
	};

	/**
	 * Add filters to the column header row
	 */
	private static Function<SheetContainer, SheetContainer> attachFilters = (SheetContainer sheetContainer) -> {
		Sheet sheet = sheetContainer.getSheet();

		// Get the row in which filter is to be applied
		int filterRow = sheetContainer.getHeading().equals("") ? 0 : 3;

		// Get the last column upto which the filters are to be applied.
		int lastColumn = sheet.getRow(sheet.getLastRowNum()).getLastCellNum() - 1;

		sheet.setAutoFilter(new CellRangeAddress(filterRow, // 1st row
				filterRow, // Last row
				0, // 1st cell
				lastColumn // Last cell
		));

		return sheetContainer;
	};

	/**
	 * This writes the current Workbook to an actual location.
	 * 
	 * @param fullPath
	 * @return The generated file.
	 */
	private static File writeToFile(String path, String fileName) {
		try {
			String fullPath = Paths.get(path, fileName).toString();
			if (fullPath == null || fullPath.equals("")) {
				return null;
			}
			File file = new File(fullPath);
			if (file.exists()) {
				log.info("As " + fullPath + " exists, deleting the file.");
				file.delete();
			}
			if (!fullPath.endsWith(".xlsx")) {
				fullPath = fullPath.concat(".xlsx");
			}
			FileOutputStream f = new FileOutputStream(file);
			Workbook workbook = workbookContainer.getWorkbook();
			workbook.write(f);
			workbook.close();
			f.close();
			return file;
		} catch (Exception e) {
			log.error("Write to workbook failed : " + e.getMessage());
			return null;
		} finally {
			cleanupFiles();
		}
	}

	/**
	 * As the name suggests, it converts a "camelCasedString" to a human readable
	 * non-"Camel Cased String".
	 * 
	 * @param camelCaseString
	 * @return Simple Readable String
	 */
	private static String parseCamelCase(String camelCaseString) {
		if (camelCaseString == null) {
			return "";
		} else {
			return StringUtil.capitalize(String.join(" ", StringUtils.splitByCharacterTypeCamelCase(camelCaseString)));
		}
	}

	/**
	 * Deletes file from the provided directory, having the mentioned name which are
	 * older than the passed no of hours.
	 * 
	 * @param fromDirectory
	 * @param havingName
	 * @param olderThan
	 */
	private static void cleanupFiles() {
		FileUtil.deleteFileOlderThanDays(BACKUP_STORAGE_PATH, BACKUP_FILE_RETENTION_DAYS);
	}

}