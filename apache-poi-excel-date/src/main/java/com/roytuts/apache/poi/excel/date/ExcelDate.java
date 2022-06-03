package com.roytuts.apache.poi.excel.date;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDate {

	public static void main(String[] args) {
		final String fileName = "excel-date.xlsx";// "excel-date.xls";
		createExcel(fileName);
	}

	public static void createExcel(final String fileName) {
		// get the file extension
		String ext = ".xlsx";
		if (fileName != null) {
			int len = fileName.trim().lastIndexOf(".");
			ext = fileName.trim().substring(len);
		}

		Workbook workbook = null;

		// based on file extension create Workbook object
		if (".xls".equalsIgnoreCase(ext)) {
			workbook = new HSSFWorkbook();
		} else if (".xlsx".equalsIgnoreCase(ext)) {
			workbook = new XSSFWorkbook();
		}

		// create Sheet object
		// sheet name must not exceed 31 characters
		// the name must not contain 0x0000, 0x0003, colon(:), backslash(\),
		// asterisk(*), question mark(?), forward slash(/), opening square
		// bracket([), closing square bracket(])
		Sheet sheet = workbook.createSheet("my_sheet");

		// Create first row. Rows are 0 based.
		Row row = sheet.createRow((short) 0);

		// Create a cell
		// put a value in cell.
		// default Date
		row.createCell(0).setCellValue(new Date());

		// style Date
		CreationHelper createHelper = workbook.getCreationHelper();
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("MM/dd/yyyy hh:mm:ss"));

		Cell cell = row.createCell(1);
		cell.setCellValue(new Date());
		cell.setCellStyle(cellStyle);

		// set date as java.util.Calendar
		CellStyle cellStyle2 = workbook.createCellStyle();
		cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("MMM/dd/yyyy hh:mm:ss"));
		cell = row.createCell(2);
		cell.setCellValue(Calendar.getInstance());
		cell.setCellStyle(cellStyle2);

		// set date as Java 8
		CellStyle cellStyle3 = workbook.createCellStyle();
		cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("MMM/dd/yyyy hh:mm:ss"));
		cell = row.createCell(3);
		cell.setCellValue(LocalDateTime.now());
		cell.setCellStyle(cellStyle3);

		CellStyle cellStyle4 = workbook.createCellStyle();
		cellStyle4.setDataFormat(createHelper.createDataFormat().getFormat("MMM/dd/yyyy"));
		cell = row.createCell(4);
		cell.setCellValue(LocalDate.now());
		cell.setCellStyle(cellStyle4);

		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream(fileName);
			workbook.write(fileOut);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				fileOut.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

}
