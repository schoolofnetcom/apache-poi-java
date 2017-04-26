package com.schoolofnet.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Word {

	public static void main(String[] args) throws IOException {
    	XWPFDocument document = new XWPFDocument();

    	String strs[] = new String[]{ "Lorem Ipsum é simplesmente uma simulação de texto da indústria tipográfica e de impressos, e vem sendo utilizado desde o século XVI.", "Luiz", "Wesley", "Lorem Ipsum é simplesmente uma simulação de texto da indústria tipográfica e de impressos, e vem sendo utilizado desde o século XVI." };
    	
    	List list = Arrays.asList(strs);
    	
    	list.forEach(data -> {
        	XWPFParagraph paragraph = document.createParagraph();
        	XWPFRun run = paragraph.createRun();
        	run.setText(data.toString());
    	});
    	
    	XWPFTable table = document.createTable();
    	XWPFTableRow tableRow = table.getRow(0);
    	tableRow.getCell(0).setText("Name");
    	tableRow.addNewTableCell().setText("Age");
    	tableRow.addNewTableCell().setText("Course");

    	XWPFTableRow tableRow2 = table.createRow();
    	tableRow2.getCell(0).setText("Leonan");
    	tableRow2.getCell(1).setText("21");
    	tableRow2.getCell(2).setText("Apache POI - Java");

    	
    	FileOutputStream out = new FileOutputStream(new File("document.docx"));    	
    	document.write(out);
    	out.close();
    	
    	System.out.print("Created");		
	}
}
