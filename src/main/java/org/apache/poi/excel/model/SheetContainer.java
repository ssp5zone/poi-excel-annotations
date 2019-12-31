package org.apache.poi.excel.model;

import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * A simple POJO that holds sheet level data.
 * 
 * @author parihsau
 */
public class SheetContainer {
	private Sheet sheet;
	private List<?> data;
	private String heading = "";

	public void setSheet(Sheet sheet) {
		this.sheet = sheet;
	}

	public void setData(List<?> data) {
		this.data = data;
	}

	public void setHeading(String heading) {
		this.heading = heading;
	}

	public Sheet getSheet() {
		return this.sheet;
	}

	public List<?> getData() {
		return this.data;
	}

	public String getHeading() {
		return this.heading;
	}
}
