package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class B {

	public static void main(String[] args) throws IOException {
		
		File loc = new File("C:\\Users\\krish\\eclipse-workspace\\Bharath\\MavenFirstProject\\target\\TEST.xlsx");
		
		FileInputStream Stream = new FileInputStream(loc);
		
		Workbook w = new XSSFWorkbook(Stream);
		Sheet s = w.getSheet("sheet1");
		
		Row r = s.getRow(1);
		Cell c = r.getCell(0);
		int type = c.getCellType();
		System.out.println(type);
			}
	
}
