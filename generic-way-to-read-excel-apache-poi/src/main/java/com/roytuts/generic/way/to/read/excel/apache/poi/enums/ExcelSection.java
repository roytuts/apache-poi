package com.roytuts.generic.way.to.read.excel.apache.poi.enums;

public enum ExcelSection {

	ORDERS("Order"), //
	PROFIT("Profit");

	final String typeValue;

	private ExcelSection(final String typeValue) {
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
