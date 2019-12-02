package com.roytuts.word.apache.poi;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class WordDocx {

	public static void main(String[] args) {
		createWordDocument("WordDocx.docx");
	}

	public static void createWordDocument(final String outputFileName) {
		// create a document
		XWPFDocument doc = new XWPFDocument();

		// create a paragraph with justify alignment
		XWPFParagraph p1 = doc.createParagraph();

		// first line indentation in the paragraph
		p1.setFirstLineIndent(400);

		// justify alignment
		p1.setAlignment(ParagraphAlignment.DISTRIBUTE);

		// wrap words
		p1.setWordWrapped(true);

		// insert page break after this paragraph
		// p1.setPageBreak(true);
		// XWPFRun object defines a region of text with a common set of
		// properties
		XWPFRun r1 = p1.createRun();
		String t1 = "Sample Paragraph Post. This is a sample Paragraph post. Sample Paragraph text is being cut and pasted again and again. This is a sample Paragraph post. peru-duellmans-poison-dart-frog."
				+ " Sample Paragraph Post. This is a sample Paragraph post. Sample Paragraph text is being cut and pasted again and again. This is a sample Paragraph post. peru-duellmans-poison-dart-frog.";
		r1.setText(t1);

		// create a paragraph with left alignment
		XWPFParagraph p2 = doc.createParagraph();

		// first line indentation in the paragraph
		p2.setFirstLineIndent(400);

		// left alignment
		p2.setAlignment(ParagraphAlignment.LEFT);

		// wrap words
		p2.setWordWrapped(true);

		// insert page break after this paragraph
		// p2.setPageBreak(true);
		// XWPFRun object defines a region of text with a common set of
		// properties
		XWPFRun r2 = p2.createRun();
		String t2 = "Sample Paragraph Post. This is a sample Paragraph post. Sample Paragraph text is being cut and pasted again and again. This is a sample Paragraph post. peru-duellmans-poison-dart-frog."
				+ " Sample Paragraph Post. This is a sample Paragraph post. Sample Paragraph text is being cut and pasted again and again. This is a sample Paragraph post. peru-duellmans-poison-dart-frog.";
		r2.setText(t2);

		// create a paragraph with ITALIC text
		XWPFParagraph p3 = doc.createParagraph();

		// left alignment
		p3.setAlignment(ParagraphAlignment.LEFT);

		// wrap words
		p3.setWordWrapped(true);

		// XWPFRun object defines a region of text with a common set of
		// properties
		XWPFRun r3 = p3.createRun();
		String t3 = "Sample Paragraph Post. This is a sample Paragraph post. Sample Paragraph text is being cut and pasted again and again. This is a sample Paragraph post. peru-duellmans-poison-dart-frog.";
		r3.setText(t3);

		// make text italic
		r3.setItalic(true);

		// create a paragraph with BOLD text
		XWPFParagraph p4 = doc.createParagraph();

		// left alignment
		p4.setAlignment(ParagraphAlignment.LEFT);

		// wrap words
		p4.setWordWrapped(true);

		// XWPFRun object defines a region of text with a common set of
		// properties
		XWPFRun r4 = p4.createRun();
		String t4 = "Sample Paragraph Post. This is a sample Paragraph post. Sample Paragraph text is being cut and pasted again and again. This is a sample Paragraph post. peru-duellmans-poison-dart-frog.";
		r4.setText(t4);

		// make text bold
		r4.setBold(true);

		// create a paragraph with Strike-Through text
		XWPFParagraph p5 = doc.createParagraph();

		// left alignment
		p5.setAlignment(ParagraphAlignment.LEFT);

		// wrap words
		p5.setWordWrapped(true);

		// XWPFRun object defines a region of text with a common set of
		// properties
		XWPFRun r5 = p5.createRun();
		String t5 = "Sample Paragraph Post. This is a sample Paragraph post. Sample Paragraph text is being cut and pasted again and again. This is a sample Paragraph post. peru-duellmans-poison-dart-frog.";
		r5.setText(t5);

		// make StrikeThrough
		r5.setStrikeThrough(true);

		// create a paragraph with Underlined text
		XWPFParagraph p6 = doc.createParagraph();

		// left alignment
		p6.setAlignment(ParagraphAlignment.LEFT);

		// wrap words
		p6.setWordWrapped(true);

		// XWPFRun object defines a region of text with a common set of
		// properties
		XWPFRun r6 = p6.createRun();
		String t6 = "Sample Paragraph Post. This is a sample Paragraph post. Sample Paragraph text is being cut and pasted again and again. This is a sample Paragraph post. peru-duellmans-poison-dart-frog.";
		r6.setText(t6);

		// make Underlined
		r6.setUnderline(UnderlinePatterns.SINGLE);

		// write to a docx file
		FileOutputStream fo = null;
		try {
			// create .docx file
			fo = new FileOutputStream(outputFileName);

			// write to the .docx file
			doc.write(fo);
		} catch (IOException e) {
		} finally {
			if (fo != null) {
				try {
					fo.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (doc != null) {
				try {
					doc.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

}
