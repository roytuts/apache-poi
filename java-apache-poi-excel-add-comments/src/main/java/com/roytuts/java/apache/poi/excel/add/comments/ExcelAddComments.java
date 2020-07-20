package com.roytuts.java.apache.poi.excel.add.comments;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Shape;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelAddComments {

	public static void main(String[] args) throws IOException {
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet();
		
		CreationHelper creationHelper = (XSSFCreationHelper) workbook.getCreationHelper();		

		Cell cell = sheet.createRow(1).createCell(2);
		cell.setCellValue("Cell One String Value");

		Drawing<Shape> drawing = (Drawing<Shape>) sheet.createDrawingPatriarch();
		ClientAnchor clientAnchor = drawing.createAnchor(0, 0, 0, 0, 0, 2, 7, 12);

		Comment comment = (Comment) drawing.createCellComment(clientAnchor);

		RichTextString richTextString = creationHelper.createRichTextString(
				"We can put a long comment here with \n a new line text followed by another \n new line text");

		comment.setString(richTextString);
		comment.setAuthor("Soumitra");

		cell.setCellComment(comment);
		
		Cell cellOther = sheet.createRow(3).createCell(2);
		cellOther.setCellValue("Cell Other String Value");
		
		ClientAnchor clientAnchorOther = drawing.createAnchor(0, 0, 0, 0, 0, 7, 12, 17);

		Comment commentOther = (Comment) drawing.createCellComment(clientAnchorOther);

		RichTextString richTextStringOther = creationHelper.createRichTextString(
				"Other long comment here with \n a new line text followed by another \n new line text");

		commentOther.setString(richTextStringOther);
		commentOther.setAuthor("Roy Tutorials");
		
		cellOther.setCellComment(commentOther);

		FileOutputStream out = new FileOutputStream("excel-add-comments.xlsx");
		workbook.write(out);		
		
		out.close();
		workbook.close();
	}

}
