package com.schoolofnet.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	
	public static void main(String[] args) throws IOException {
		FileInputStream input = new FileInputStream(new File("workbook.xlsx"));
		
		XSSFWorkbook workbook = new XSSFWorkbook(input);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		Iterator<Row> iterator = (Iterator<Row>) sheet.iterator();
		
		while(iterator.hasNext()) {
			Row currentRow = iterator.next();
			Iterator<Cell> cell = currentRow.iterator();
			
			while(cell.hasNext()) {
				Cell currentCell = cell.next();
				
				if (currentCell.getCellTypeEnum() == CellType.STRING) {
					System.out.println(currentCell.getStringCellValue());
				}
				if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
					System.out.println(currentCell.getNumericCellValue() + "");
				}
			}
		}
	}
}	
