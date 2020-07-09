package com.roytuts.java.apache.poi.excel.filter.number.column;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCustomFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilterColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STFilterOperator;

public class ExcelCustomFiltersOnNumberApp {

	public static void main(String[] args) throws InvalidFormatException, IOException {
		ExcelCustomFiltersOnNumberApp customFiltersApp = new ExcelCustomFiltersOnNumberApp();

		customFiltersApp.filterEquals("sample.xlsx", 0, 0, 9);
		customFiltersApp.filterDoesNotEqual("sample.xlsx", 0, 0, 75);

		customFiltersApp.filterGreaterThan("sample.xlsx", 0, 0, 768);
		customFiltersApp.filterGreaterThanOrEqualTo("sample.xlsx", 0, 0, 765);

		customFiltersApp.filterLessThan("sample.xlsx", 0, 0, 768);
		customFiltersApp.filterLessThanOrEqualTo("sample.xlsx", 0, 0, 765);

		customFiltersApp.filterBetween("sample.xlsx", 0, 0, 500, 700);

		customFiltersApp.filterTop10("sample.xlsx", 0, 0);
		customFiltersApp.filterAboveAverage("sample.xlsx", 0, 0);
		customFiltersApp.filterBelowAverage("sample.xlsx", 0, 0);
	}

	public void filterEquals(String fileName, int sheetIndex, int columnNumber, Integer search)
			throws IOException, InvalidFormatException {
		Workbook workbook = new XSSFWorkbook(new File(fileName));
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);

		CTAutoFilter autoFilter = sheet.getCTWorksheet().addNewAutoFilter();

		CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
		filterColumn.setColId(columnNumber);

		CTFilter ctFilter = filterColumn.addNewFilters().insertNewFilter(columnNumber);
		ctFilter.setVal(String.valueOf(search));

		boolean skipFirstRow = true;
		for (Row r : sheet) {
			if (skipFirstRow) {
				skipFirstRow = false;
				continue;
			}
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex() && c.getNumericCellValue() != search) {
					XSSFRow row = (XSSFRow) c.getRow();
					row.getCTRow().setHidden(true);
				}
			}
		}

		OutputStream outputStream = new FileOutputStream(
				new File(fileName.substring(0, fileName.lastIndexOf(".")) + "_equals.xlsx"));
		workbook.write(outputStream);
		outputStream.close();
	}

	public void filterDoesNotEqual(String fileName, int sheetIndex, int columnNumber, Integer search)
			throws IOException, InvalidFormatException {
		Workbook workbook = new XSSFWorkbook(new File(fileName));
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);

		CTAutoFilter autoFilter = sheet.getCTWorksheet().addNewAutoFilter();

		CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
		filterColumn.setColId(columnNumber);

		CTFilter ctFilter = filterColumn.addNewFilters().insertNewFilter(columnNumber);
		ctFilter.setVal(String.valueOf(search));

		boolean skipFirstRow = true;
		for (Row r : sheet) {
			if (skipFirstRow) {
				skipFirstRow = false;
				continue;
			}
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex() && c.getNumericCellValue() == search) {
					XSSFRow row = (XSSFRow) c.getRow();
					row.getCTRow().setHidden(true);
				}
			}
		}

		OutputStream outputStream = new FileOutputStream(
				new File(fileName.substring(0, fileName.lastIndexOf(".")) + "_does_not_equal.xlsx"));
		workbook.write(outputStream);
		outputStream.close();
	}

	public void filterGreaterThan(String fileName, int sheetIndex, int columnNumber, Integer search)
			throws IOException, InvalidFormatException {
		Workbook workbook = new XSSFWorkbook(new File(fileName));
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);

		CTAutoFilter autoFilter = sheet.getCTWorksheet().addNewAutoFilter();

		CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
		filterColumn.setColId(columnNumber);

		CTFilter ctFilter = filterColumn.addNewFilters().insertNewFilter(columnNumber);
		ctFilter.setVal(String.valueOf(search));

		boolean skipFirstRow = true;
		for (Row r : sheet) {
			if (skipFirstRow) {
				skipFirstRow = false;
				continue;
			}
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex() && c.getNumericCellValue() <= search) {
					XSSFRow row = (XSSFRow) c.getRow();
					row.getCTRow().setHidden(true);
				}
			}
		}

		OutputStream outputStream = new FileOutputStream(
				new File(fileName.substring(0, fileName.lastIndexOf(".")) + "_greater_than.xlsx"));
		workbook.write(outputStream);
		outputStream.close();
	}

	public void filterGreaterThanOrEqualTo(String fileName, int sheetIndex, int columnNumber, Integer search)
			throws IOException, InvalidFormatException {
		Workbook workbook = new XSSFWorkbook(new File(fileName));
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);

		CTAutoFilter autoFilter = sheet.getCTWorksheet().addNewAutoFilter();

		CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
		filterColumn.setColId(columnNumber);

		CTFilter ctFilter = filterColumn.addNewFilters().insertNewFilter(columnNumber);
		ctFilter.setVal(String.valueOf(search));

		boolean skipFirstRow = true;
		for (Row r : sheet) {
			if (skipFirstRow) {
				skipFirstRow = false;
				continue;
			}
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex() && c.getNumericCellValue() < search) {
					XSSFRow row = (XSSFRow) c.getRow();
					row.getCTRow().setHidden(true);
				}
			}
		}

		OutputStream outputStream = new FileOutputStream(
				new File(fileName.substring(0, fileName.lastIndexOf(".")) + "_greater_than_or_equal_to.xlsx"));
		workbook.write(outputStream);
		outputStream.close();
	}

	public void filterLessThan(String fileName, int sheetIndex, int columnNumber, Integer search)
			throws IOException, InvalidFormatException {
		Workbook workbook = new XSSFWorkbook(new File(fileName));
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);

		CTAutoFilter autoFilter = sheet.getCTWorksheet().addNewAutoFilter();

		CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
		filterColumn.setColId(columnNumber);

		CTFilter ctFilter = filterColumn.addNewFilters().insertNewFilter(columnNumber);
		ctFilter.setVal(String.valueOf(search));

		boolean skipFirstRow = true;
		for (Row r : sheet) {
			if (skipFirstRow) {
				skipFirstRow = false;
				continue;
			}
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex() && c.getNumericCellValue() >= search) {
					XSSFRow row = (XSSFRow) c.getRow();
					row.getCTRow().setHidden(true);
				}
			}
		}

		OutputStream outputStream = new FileOutputStream(
				new File(fileName.substring(0, fileName.lastIndexOf(".")) + "_less_than.xlsx"));
		workbook.write(outputStream);
		outputStream.close();
	}

	public void filterLessThanOrEqualTo(String fileName, int sheetIndex, int columnNumber, Integer search)
			throws IOException, InvalidFormatException {
		Workbook workbook = new XSSFWorkbook(new File(fileName));
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);

		CTAutoFilter autoFilter = sheet.getCTWorksheet().addNewAutoFilter();

		CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
		filterColumn.setColId(columnNumber);

		CTFilter ctFilter = filterColumn.addNewFilters().insertNewFilter(columnNumber);
		ctFilter.setVal(String.valueOf(search));

		boolean skipFirstRow = true;
		for (Row r : sheet) {
			if (skipFirstRow) {
				skipFirstRow = false;
				continue;
			}
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex() && c.getNumericCellValue() > search) {
					XSSFRow row = (XSSFRow) c.getRow();
					row.getCTRow().setHidden(true);
				}
			}
		}

		OutputStream outputStream = new FileOutputStream(
				new File(fileName.substring(0, fileName.lastIndexOf(".")) + "_less_than_or_equal_to.xlsx"));
		workbook.write(outputStream);
		outputStream.close();
	}

	public void filterBetween(String fileName, int sheetIndex, int columnNumber, Integer from, Integer to)
			throws IOException, InvalidFormatException {
		Workbook workbook = new XSSFWorkbook(new File(fileName));
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);

		CTAutoFilter autoFilter = sheet.getCTWorksheet().addNewAutoFilter();

		CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
		filterColumn.setColId(columnNumber);

		CTCustomFilter ctCustomFilter = filterColumn.addNewCustomFilters().insertNewCustomFilter(columnNumber);

		ctCustomFilter.setOperator(STFilterOperator.GREATER_THAN);
		ctCustomFilter.setVal(String.valueOf(from));

		ctCustomFilter.setOperator(STFilterOperator.LESS_THAN);
		ctCustomFilter.setVal(String.valueOf(to));

		boolean skipFirstRow = true;
		for (Row r : sheet) {
			if (skipFirstRow) {
				skipFirstRow = false;
				continue;
			}
			for (Cell c : r) {
				double val = c.getNumericCellValue();
				if (columnNumber == c.getColumnIndex() && (val < from || val > to)) {
					XSSFRow row = (XSSFRow) c.getRow();
					row.getCTRow().setHidden(true);
				}
			}
		}

		OutputStream outputStream = new FileOutputStream(
				new File(fileName.substring(0, fileName.lastIndexOf(".")) + "_between.xlsx"));
		workbook.write(outputStream);
		outputStream.close();
	}

	public void filterTop10(String fileName, int sheetIndex, int columnNumber)
			throws IOException, InvalidFormatException {
		Workbook workbook = new XSSFWorkbook(new File(fileName));
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);

		CTAutoFilter autoFilter = sheet.getCTWorksheet().addNewAutoFilter();

		CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
		filterColumn.setColId(columnNumber);

		List<Integer> nums = new ArrayList<>();

		boolean skipFirstRow = true;
		for (Row r : sheet) {
			if (skipFirstRow) {
				skipFirstRow = false;
				continue;
			}
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex()) {
					nums.add((int) c.getNumericCellValue());
				}
			}
		}

		Integer[] numbers = nums.toArray(new Integer[nums.size()]);

		Arrays.sort(numbers, Collections.reverseOrder());

		int[] arr = new int[10];
		for (int i = 0; i < 10; i++) {
			arr[i] = numbers[i];
		}

		int arrSize = arr.length;
		skipFirstRow = true;
		for (Row r : sheet) {
			if (skipFirstRow) {
				skipFirstRow = false;
				continue;
			}
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex()) {
					boolean hide = true;
					int val = (int) c.getNumericCellValue();
					for (int i = 0; i < arrSize; i++) {
						if (arr[i] == val) {
							hide = false;
							break;
						}
					}
					if (hide) {
						XSSFRow row = (XSSFRow) c.getRow();
						row.getCTRow().setHidden(hide);
					}
				}
			}
		}

		OutputStream outputStream = new FileOutputStream(
				new File(fileName.substring(0, fileName.lastIndexOf(".")) + "_top_10.xlsx"));
		workbook.write(outputStream);
		outputStream.close();
	}

	public void filterAboveAverage(String fileName, int sheetIndex, int columnNumber)
			throws IOException, InvalidFormatException {
		Workbook workbook = new XSSFWorkbook(new File(fileName));
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);

		CTAutoFilter autoFilter = sheet.getCTWorksheet().addNewAutoFilter();

		CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
		filterColumn.setColId(columnNumber);

		List<Integer> nums = new ArrayList<>();

		boolean skipFirstRow = true;
		for (Row r : sheet) {
			if (skipFirstRow) {
				skipFirstRow = false;
				continue;
			}
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex()) {
					nums.add((int) c.getNumericCellValue());
				}
			}
		}

		int sum = 0;
		for (Integer integer : nums) {
			sum += integer;
		}

		double avg = sum / nums.size();

		int size = nums.size();
		skipFirstRow = true;
		for (Row r : sheet) {
			if (skipFirstRow) {
				skipFirstRow = false;
				continue;
			}
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex()) {
					boolean hide = true;
					double val = (int) c.getNumericCellValue();
					for (int i = 0; i < size; i++) {
						if (val > avg) {
							hide = false;
							break;
						}
					}
					if (hide) {
						XSSFRow row = (XSSFRow) c.getRow();
						row.getCTRow().setHidden(hide);
					}
				}
			}
		}

		OutputStream outputStream = new FileOutputStream(
				new File(fileName.substring(0, fileName.lastIndexOf(".")) + "_above_average.xlsx"));
		workbook.write(outputStream);
		outputStream.close();
	}

	public void filterBelowAverage(String fileName, int sheetIndex, int columnNumber)
			throws IOException, InvalidFormatException {
		Workbook workbook = new XSSFWorkbook(new File(fileName));
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);

		CTAutoFilter autoFilter = sheet.getCTWorksheet().addNewAutoFilter();

		CTFilterColumn filterColumn = autoFilter.addNewFilterColumn();
		filterColumn.setColId(columnNumber);

		List<Integer> nums = new ArrayList<>();

		boolean skipFirstRow = true;
		for (Row r : sheet) {
			if (skipFirstRow) {
				skipFirstRow = false;
				continue;
			}
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex()) {
					nums.add((int) c.getNumericCellValue());
				}
			}
		}

		int sum = 0;
		for (Integer integer : nums) {
			sum += integer;
		}

		double avg = sum / nums.size();

		int size = nums.size();
		skipFirstRow = true;
		for (Row r : sheet) {
			if (skipFirstRow) {
				skipFirstRow = false;
				continue;
			}
			for (Cell c : r) {
				if (columnNumber == c.getColumnIndex()) {
					boolean hide = true;
					double val = (int) c.getNumericCellValue();
					for (int i = 0; i < size; i++) {
						if (val < avg) {
							hide = false;
							break;
						}
					}
					if (hide) {
						XSSFRow row = (XSSFRow) c.getRow();
						row.getCTRow().setHidden(hide);
					}
				}
			}
		}

		OutputStream outputStream = new FileOutputStream(
				new File(fileName.substring(0, fileName.lastIndexOf(".")) + "_below_average.xlsx"));
		workbook.write(outputStream);
		outputStream.close();
	}

}
