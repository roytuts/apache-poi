package com.roytuts.generic.way.to.read.excel.apache.poi.model;

public class ExcelField {

	private String excelHeader;
	private int excelIndex;
	private String excelColType;
	private String excelValue;
	private String pojoAttribute;

	public String getExcelHeader() {
		return excelHeader;
	}

	public void setExcelHeader(String excelHeader) {
		this.excelHeader = excelHeader;
	}

	public int getExcelIndex() {
		return excelIndex;
	}

	public void setExcelIndex(int excelIndex) {
		this.excelIndex = excelIndex;
	}

	public String getExcelColType() {
		return excelColType;
	}

	public void setExcelColType(String excelColType) {
		this.excelColType = excelColType;
	}

	public String getExcelValue() {
		return excelValue;
	}

	public void setExcelValue(String excelValue) {
		this.excelValue = excelValue;
	}

	public String getPojoAttribute() {
		return pojoAttribute;
	}

	public void setPojoAttribute(String pojoAttribute) {
		this.pojoAttribute = pojoAttribute;
	}

}
