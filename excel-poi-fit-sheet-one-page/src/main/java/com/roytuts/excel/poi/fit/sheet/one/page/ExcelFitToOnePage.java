package com.roytuts.excel.poi.fit.sheet.one.page;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFitToOnePage {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		// InputStream fis = new FileInputStream(new
		// File("C:/jee_workspace/sample.xlsx"));

		InputStream fis = new FileInputStream(new File("C:/jee_workspace/sample-with-image.xlsx"));

		Workbook wb = WorkbookFactory.create(fis);

		Sheet sheet = wb.getSheetAt(0);

		PrintSetup ps = sheet.getPrintSetup();

		sheet.setFitToPage(true);
		sheet.setAutobreaks(true);

		ps.setFitWidth((short) 1);
		ps.setFitHeight((short) 1);

		//OutputStream fileOut = new FileOutputStream("C:/jee_workspace/sample-fit-one-page.xlsx");
		
		OutputStream fileOut = new FileOutputStream("C:/jee_workspace/sample-with-image-fit-one-page.xlsx");
		
		wb.write(fileOut);

		wb.close();

	}

}
