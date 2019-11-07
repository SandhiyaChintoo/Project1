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

public class A {

	public static void main(String[] args) throws IOException {
		
		File loc = new File("C:\\Users\\krish\\eclipse-workspace\\Bharath\\MavenFirstProject\\target\\TEST.xlsx");
		
		FileInputStream Stream = new FileInputStream(loc);
		
		Workbook w = new XSSFWorkbook(Stream);
		Sheet s = w.getSheet("sheet1");
		
		for(int i=0; i<s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			
			for(int j=0; j<r.getPhysicalNumberOfCells(); j++) {
				
		Cell c = r.getCell(j);
		
		System.out.println(c);
			}
		}
		
			
	}
}
