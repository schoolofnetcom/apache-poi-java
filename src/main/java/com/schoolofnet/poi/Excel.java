package com.schoolofnet.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javafx.scene.control.Cell;

public class Excel {
	public static void main(String[] args) throws IOException {
    	XSSFWorkbook workbook = new XSSFWorkbook();
    	XSSFSheet sheet = workbook.createSheet("Sheet 1");
    	
    	XSSFRow row;
    	
    	Map<String, Object[]> infos = new TreeMap<String, Object[]>();

    	// Header of my sheet	       0       1        2
    	infos.put("0", new Object[]{ "Name", "Age", "Course" });

    	// Datas of my sheet            0       1        2
    	infos.put("1", new Object[]{ "Leonan", "21", "Apache POI - Java" });
    	infos.put("2", new Object[]{ "Wesley", "32", "Python Django" });
    	infos.put("3", new Object[]{ "Erik", "32", "Angular2" });
    	infos.put("4", new Object[]{ "Luiz", "30", "Laravel" });
    	
    	for(Map.Entry<String, Object[]> entry : infos.entrySet()) {
    		String key = entry.getKey();
    		Object[] data = entry.getValue();
    		
    		row = sheet.createRow(Integer.parseInt(key));
    		
    		int cellIndex = 0;
    		for (Object obj : data) {
    			XSSFCell cell = row.createCell(cellIndex++);
    			cell.setCellValue((String) obj);
    		}
    	}
    	
    	FileOutputStream out = new FileOutputStream(new File("workbook.xlsx"));    	
    	workbook.write(out);
    	out.close();
    	System.out.print("Created");
	}

}
