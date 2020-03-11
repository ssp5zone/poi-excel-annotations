# Apache POI Excel Writer & Annotations

A smart Java utility to write **any** `List` of Java Objects to an Excel/CSV file with just one line of code.

## Usage

1. Get the jar. (Download JAR | Github Release | Maven Repo)
2. Add this as a dependency in your project.
3. Use it as below.

``` Java
// Your existing code
List<Employee> employees = ...;
List<Department> departments = ...;
// ... other lists

// The one line code
ExcelWriter.write("/path/to/write", "MyFile.xlsx", employees, departments, ...);
```

**That's it!** 

This would create an Excel file called _MyFile.xlsx_ with 1st sheet as Employees, next sheet as Departments and so on. 

## Sample

You can find some of the sample generated files [here](/src/test/resources/output/). 

## Details

### Problem:

One of the most widely used Java library for reading and writing Excel files is: Apache POI. Though it provides many options for customizing the excel file, it lacks a very basic feature - binding an existing POJO to read and write the data. There are some off-the-shelf open source solutions for binding POJO while reading files, but none exists for writing.

Currently for writing an excel file, the developer has to loop through - every - single - cell of the excel, get the data, place it in the cell and then style it to the correct format. Imagine an excel sheet with 80 columns and 20,000 rows. This would end up in a code that has at-least (80 -160) lines of code to set the data, another 80 to just style the cell, 80-160 lines to write the column names and whatever extra code needed to parse the data to correct format (like Date). On top of this, the code would be inside a couple of for loops that would run 20,000 times!. Result,


> ~400 lines of raw code.

The worse part, all this just for one sheet of one excel file. For each sheet of each excel, you create one such processing file.
So considering 3 such sheets, there'd be total of

> ~1200 lines of code!!

No doubt you can optimize the above numbers. That's what this library ultimately does. 

### Solution:


On a simple inspection one would note that parts of code in these kinds of files remain common.

Moreover, there are other parts which can be deduced on-the-go based on the field names and types.


#### 1. Generic (ANY) POJO List to Excel Writer


A properly written program usually has a data object model. A data model is nothing but a plain old java object that merely contains a set of attributes it holds. This encapsulation contains the entirety of data when passing it between systems or processes (like - to and from Data Base or Files, HTTP or REST calls, even during internal processing).

For instance,

``` Java
// Employee POJO class to represent entity Employee 
public class Employee 
{ 
    public String name;
	public String employer;
	public int age;
	public long salary;
	public Date dateOfBirth;
	public boolean retained;
	public char grade;
	public byte rank;
	public double latitude;
	public double longitude;
	public short height;
	public float weight;
} 
```

Now, most project modules would have entities just like these that holds a list of data which needs to be written to Excel file(s) as it is. 

For this, we created a common utility that can be invoked as,

``` Java
ExcelWriter.write("/path/to/write", "MyFile.xlsx", employeeList);
```

That's it!! It automatically creates an Excel for you at specified location.

![Simple Sheet Image](/docs/NonAnnotated.png "Directly generated Excel Sheet")

The utility creates the suitable columns based on the property type _(Eg. Numeric, Date, Precision etc)_ and applies filter to them.


If you want multiple such sheets that need to be written to same excel,

``` Java
ExcelWriter.write("/path/to/write", "MyFile.xlsx", employeeList, departmentList, salaryBreakDown, ...); 
```

> The ~1200 lines reduced to just 1.


#### 2. @Excel Annotations (Optional)

Now, what if we want to customize the data instead of presenting as-is. Say, I doan't want to display all the columns or I want to provide custom styling or column names.

For this we created a couple of annotations,

1. **@ExcelSheet**
2. **@ExcelCell**


Applying these 2 to the above example,

``` Java
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
    
    public String thisWillNotAppearInExcel;
    public String thisToo;
```

Now, the only the attributes with `@ExcelCell` annotation would appear in the sheet.

The invocation call remains the same.

The generated excel,

![Annotated Sheet Image](/docs/Annotated.png "Excel generated using Annotations")

Please read the docs to see what each Annotation does in details.

## Contributing to Project

There is still some work yet to be done in this project. Please feel free to fork, moidfy and raise a PR for the same.

## Licence

Licenced under MIT License 2019. It is free to copy and distribute.