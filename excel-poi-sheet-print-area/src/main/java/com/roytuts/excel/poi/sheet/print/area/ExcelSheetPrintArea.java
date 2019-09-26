package com.roytuts.excel.poi.sheet.print.area;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelSheetPrintArea {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		InputStream fis = new FileInputStream(new File("C:/jee_workspace/sample.xlsx"));

		Workbook wb = WorkbookFactory.create(fis);

		wb.setPrintArea(0, 0, 4, 0, 5);

		OutputStream fileOut = new FileOutputStream("C:/jee_workspace/sample-print-area.xlsx");

		wb.write(fileOut);

		wb.close();

	}

}
