package com.roytuts.java.apache.poi.excel.filter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilterColumn;

public class ExcelCustomFiltersApp {

	public static void main(String[] args) throws InvalidFormatException, IOException {
		ExcelCustomFiltersApp customFiltersApp = new ExcelCustomFiltersApp();

		customFiltersApp.filterEquals("info.xlsx", 0, 0, "Loku");
		customFiltersApp.filterDoesNotEqual("info.xlsx", 0, 0, "Souvik");

		customFiltersApp.filterBeginsWith("info.xlsx", 0, 0, "S");
		customFiltersApp.filterEndsWith("info.xlsx", 0, 0, "t");

		customFiltersApp.filterContains("info.xlsx", 0, 0, "L");
		customFiltersApp.filterDoesNotContain("info.xlsx", 0, 0, "S");
	}

	public void filterEquals(String fileName, int sheetIndex, int columnNumber, String search)
			throws IOException, InvalidFormatException {
		Workbook workbook = new XSSFWorkbook(new File(fileName));
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);

		CTAutoFilter autoFilter = sheet.getCTWorksheet().addNewAutoFilter();

		CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
		filterColumn.setColId(columnNumber);

		CTFilter ctFilter = filterColumn.addNewFilters().insertNewFilter(columnNumber);
		ctFilter.setVal(search);

		for (Row r : sheet) {
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex() && !c.getStringCellValue().equals(search)) {
					XSSFRow row = (XSSFRow) c.getRow();
					if (0 != row.getRowNum()) {
						row.getCTRow().setHidden(true);
					}
				}
			}
		}

		OutputStream outputStream = new FileOutputStream(
				new File(fileName.substring(0, fileName.lastIndexOf(".")) + "_equals.xlsx"));
		workbook.write(outputStream);
		outputStream.close();
	}

	public void filterDoesNotEqual(String fileName, int sheetIndex, int columnNumber, String search)
			throws IOException, InvalidFormatException {
		Workbook workbook = new XSSFWorkbook(new File(fileName));
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);

		CTAutoFilter autoFilter = sheet.getCTWorksheet().addNewAutoFilter();

		CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
		filterColumn.setColId(columnNumber);

		CTFilter ctFilter = filterColumn.addNewFilters().insertNewFilter(columnNumber);
		ctFilter.setVal(search);

		for (Row r : sheet) {
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex() && c.getStringCellValue().equals(search)) {
					XSSFRow row = (XSSFRow) c.getRow();
					if (0 != row.getRowNum()) {
						row.getCTRow().setHidden(true);
					}
				}
			}
		}

		OutputStream outputStream = new FileOutputStream(
				new File(fileName.substring(0, fileName.lastIndexOf(".")) + "_does_not_equal.xlsx"));
		workbook.write(outputStream);
		outputStream.close();
	}

	public void filterBeginsWith(String fileName, int sheetIndex, int columnNumber, String search)
			throws IOException, InvalidFormatException {
		Workbook workbook = new XSSFWorkbook(new File(fileName));
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);

		CTAutoFilter autoFilter = sheet.getCTWorksheet().addNewAutoFilter();

		CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
		filterColumn.setColId(columnNumber);

		CTFilter ctFilter = filterColumn.addNewFilters().insertNewFilter(columnNumber);
		ctFilter.setVal(search);

		for (Row r : sheet) {
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex() && !c.getStringCellValue().startsWith(search)) {
					XSSFRow row = (XSSFRow) c.getRow();
					if (0 != row.getRowNum()) {
						row.getCTRow().setHidden(true);
					}
				}
			}
		}

		OutputStream outputStream = new FileOutputStream(
				new File(fileName.substring(0, fileName.lastIndexOf(".")) + "_begins_with.xlsx"));
		workbook.write(outputStream);
		outputStream.close();
	}

	public void filterEndsWith(String fileName, int sheetIndex, int columnNumber, String search)
			throws IOException, InvalidFormatException {
		Workbook workbook = new XSSFWorkbook(new File(fileName));
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);

		CTAutoFilter autoFilter = sheet.getCTWorksheet().addNewAutoFilter();

		CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
		filterColumn.setColId(columnNumber);

		CTFilter ctFilter = filterColumn.addNewFilters().insertNewFilter(columnNumber);
		ctFilter.setVal(search);

		for (Row r : sheet) {
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex() && !c.getStringCellValue().endsWith(search)) {
					XSSFRow row = (XSSFRow) c.getRow();
					if (0 != row.getRowNum()) {
						row.getCTRow().setHidden(true);
					}
				}
			}
		}

		OutputStream outputStream = new FileOutputStream(
				new File(fileName.substring(0, fileName.lastIndexOf(".")) + "_ends_with.xlsx"));
		workbook.write(outputStream);
		outputStream.close();
	}

	public void filterContains(String fileName, int sheetIndex, int columnNumber, String search)
			throws IOException, InvalidFormatException {
		Workbook workbook = new XSSFWorkbook(new File(fileName));
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);

		CTAutoFilter autoFilter = sheet.getCTWorksheet().addNewAutoFilter();

		CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
		filterColumn.setColId(columnNumber);

		CTFilter ctFilter = filterColumn.addNewFilters().insertNewFilter(columnNumber);
		ctFilter.setVal(search);

		for (Row r : sheet) {
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex() && !c.getStringCellValue().contains(search)) {
					XSSFRow row = (XSSFRow) c.getRow();
					if (0 != row.getRowNum()) {
						row.getCTRow().setHidden(true);
					}
				}
			}
		}

		OutputStream outputStream = new FileOutputStream(
				new File(fileName.substring(0, fileName.lastIndexOf(".")) + "_contains.xlsx"));
		workbook.write(outputStream);
		outputStream.close();
	}

	public void filterDoesNotContain(String fileName, int sheetIndex, int columnNumber, String search)
			throws IOException, InvalidFormatException {
		Workbook workbook = new XSSFWorkbook(new File(fileName));
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);

		CTAutoFilter autoFilter = sheet.getCTWorksheet().addNewAutoFilter();

		CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
		filterColumn.setColId(columnNumber);

		CTFilter ctFilter = filterColumn.addNewFilters().insertNewFilter(columnNumber);
		ctFilter.setVal(search);

		for (Row r : sheet) {
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex() && c.getStringCellValue().contains(search)) {
					XSSFRow row = (XSSFRow) c.getRow();
					if (0 != row.getRowNum()) {
						row.getCTRow().setHidden(true);
					}
				}
			}
		}

		OutputStream outputStream = new FileOutputStream(
				new File(fileName.substring(0, fileName.lastIndexOf(".")) + "_does_not_contain.xlsx"));
		workbook.write(outputStream);
		outputStream.close();
	}

}
