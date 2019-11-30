package com.roytuts.generic.way.to.read.excel.apache.poi.enums;

public enum FieldType {

	DOUBLE("Double"), //
	INTEGER("Integer"), //
	STRING("String"), //
	DATE("Date");
	
	final String typeValue;
	
	private FieldType(final String typeValue) {
		this.typeValue = typeValue;
	}
	
	public String getName() {
		return name();
	}
	
	public String getValue() {
		return typeValue;
	}
	
	@Override
	public String toString() {
		return name();
	}
	
}
