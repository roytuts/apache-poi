package com.roytuts.apache.poi.excel.add.images;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AddImagesToExcelFile {

	public static void main(String[] args) throws IOException {
		addImages();
	}

	public static void addImages() throws IOException {
		// create a new workbook
		Workbook wb = new XSSFWorkbook(); // or new HSSFWorkbook();

		// add jpg/jpeg and png formats picture data to this workbook.
		InputStream is = new FileInputStream("sample.jpg");
		byte[] bytes = IOUtils.toByteArray(is);
		int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
		is.close();

		CreationHelper helper = wb.getCreationHelper();

		// create sheet
		Sheet sheet = wb.createSheet();

		// Create the drawing patriarch. This is the top level container for all shapes.
		Drawing<?> drawing = sheet.createDrawingPatriarch();

		// add a picture shape
		ClientAnchor anchor = helper.createClientAnchor();

		// set top-left corner of the picture,
		// subsequent call of Picture#resize() will operate relative to it
		anchor.setCol1(3);
		anchor.setRow1(2);
		Picture pict = drawing.createPicture(anchor, pictureIdx);

		// auto-size picture relative to its top-left corner
		pict.resize();

		is = new FileInputStream("sample.png");
		bytes = IOUtils.toByteArray(is);
		pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
		is.close();

		anchor = helper.createClientAnchor();
		anchor.setCol1(3);
		anchor.setRow1(14);
		pict = drawing.createPicture(anchor, pictureIdx);

		// auto-size picture relative to its top-left corner
		pict.resize();

		// save workbook
		String fileName = "excel_image.xls";

		if (wb instanceof XSSFWorkbook) {
			fileName += "x";
		}

		try (OutputStream fileOut = new FileOutputStream(new File(fileName))) {
			wb.write(fileOut);
		}

		wb.close();
	}

}
