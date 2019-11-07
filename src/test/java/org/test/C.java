package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class C {

	public static void main(String[] args) throws IOException {
		
		File loc = new File("C:\\Users\\krish\\eclipse-workspace\\Bharath\\MavenFirstProject\\target\\TEST.xlsx");
		
		FileInputStream Stream = new FileInputStream(loc);
		
		Workbook w = new XSSFWorkbook(Stream);
		Sheet s = w.getSheet("sheet1");
		
		for(int i=0; i<s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			for(int j=0; j<r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(j);
				int type = c.getCellType();
				
				if(type==1) {
					String stringCellValue = c.getStringCellValue();
					System.out.println(stringCellValue);
				}
				else if(type==0) {
					if(DateUtil.isCellDateFormatted(c)) {
						
						Date dateCellValue = c.getDateCellValue();
						
						SimpleDateFormat sim = new SimpleDateFormat("dd-MMM-yy");
						String f = sim.format(dateCellValue);
						System.out.println(f);
						
								
					}
				
					else {
						double numericCellValue = c.getNumericCellValue();
						
						long l = (long) numericCellValue;	
						String v = String.valueOf(l);
						System.out.println(v);
					}
				}
				
			}
		}
			}
	
}
