package com.schoolofnet.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

public class ReadWord {
	
	public static void main(String[] args) throws IOException {
		FileInputStream input = new FileInputStream(new File("document.docx"));
	
		XWPFDocument document = new XWPFDocument(input);
		
		List<XWPFParagraph> paragraphs = document.getParagraphs();
				
		for (XWPFParagraph paragraph : paragraphs) {
			System.out.println(paragraph.getText());
		}
		
		System.out.println("Total of paragraphs: " +  paragraphs.size());
	}
}
