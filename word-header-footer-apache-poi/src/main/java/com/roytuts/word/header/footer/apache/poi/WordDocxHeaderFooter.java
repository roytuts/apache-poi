package com.roytuts.word.header.footer.apache.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

public class WordDocxHeaderFooter {

	public static void main(String[] args) {
		wordHeaderFooter("word_docx_header_footer.docx");
	}

	public static void wordHeaderFooter(final String wordFileName) {
		XWPFDocument doc = new XWPFDocument();

		// create a paragraph with justify alignment
		XWPFParagraph p1 = doc.createParagraph();

		// first line indentation in the paragraph
		p1.setFirstLineIndent(400);

		// justify alignment
		p1.setAlignment(ParagraphAlignment.DISTRIBUTE);

		// wrap words
		p1.setWordWrapped(true);

		// XWPFRun object defines a region of text with a common set of
		// properties
		XWPFRun r1 = p1.createRun();
		String t1 = "Paragraph 1. Sample Paragraph Post. This is a sample Paragraph post. Sample Paragraph text is being cut and pasted again and again. This is a sample Paragraph post. peru-duellmans-poison-dart-frog."
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

		// XWPFRun object defines a region of text with a common set of
		// properties
		XWPFRun r2 = p2.createRun();
		String t2 = "Paragraph 2. Sample Paragraph Post. This is a sample Paragraph post. Sample Paragraph text is being cut and pasted again and again. This is a sample Paragraph post. peru-duellmans-poison-dart-frog."
				+ " Sample Paragraph Post. This is a sample Paragraph post. Sample Paragraph text is being cut and pasted again and again. This is a sample Paragraph post. peru-duellmans-poison-dart-frog.";
		r2.setText(t2);

		XWPFParagraph[] pars;

		try {
			CTP ctP = CTP.Factory.newInstance();

			// header text
			CTText t = ctP.addNewR().addNewT();
			t.setStringValue("Sample Header Text");

			pars = new XWPFParagraph[1];
			p1 = new XWPFParagraph(ctP, doc);
			pars[0] = p1;

			XWPFHeaderFooterPolicy hfPolicy = doc.createHeaderFooterPolicy();
			hfPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, pars);

			ctP = CTP.Factory.newInstance();
			t = ctP.addNewR().addNewT();

			// footer text
			t.setStringValue("Sample Footer Text");

			pars[0] = new XWPFParagraph(ctP, doc);
			hfPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, pars);

			// write to word docx
			OutputStream os = new FileOutputStream(new File(wordFileName));
			doc.write(os);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
