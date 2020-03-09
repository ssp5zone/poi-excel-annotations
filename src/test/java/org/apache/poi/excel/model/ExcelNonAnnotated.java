package org.apache.poi.excel.model;

import java.util.Date;

public class ExcelNonAnnotated {

	public String agentName;
	public String agency;
	public int age;
	public long salary;
	public Date dateOfBirth;
	public boolean isActive;
	public char grade;
	public byte rank;
	public double latitude;
	public double longitude;
	public short height;
	public float weight;

	/**
	 * @param agentName
	 * @param agency
	 * @param age
	 * @param salary
	 * @param dateOfBirth
	 * @param isActive
	 * @param grade
	 * @param rank
	 * @param latitude
	 * @param longitude
	 * @param height
	 * @param weight
	 */
	public ExcelNonAnnotated(String agentName, String agency, int age, long salary, Date dateOfBirth, boolean isActive,
			char grade, byte rank, double latitude, double longitude, short height, float weight) {
		this.agentName = agentName;
		this.agency = agency;
		this.age = age;
		this.salary = salary;
		this.dateOfBirth = dateOfBirth;
		this.isActive = isActive;
		this.grade = grade;
		this.rank = rank;
		this.latitude = latitude;
		this.longitude = longitude;
		this.height = height;
		this.weight = weight;
	}

}
