package org.d;
import java.io.File;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Scenario4 {
public static void main(String[] args) throws IOException  {
		
		
		File loc =new File ("C:\\Users\\Sundeep\\Documents\\workspace-sts-3.9.11.RELEASE\\DATADRIVEN\\excel\\New Microsoft Excel Worksheet.xlsx");
		FileInputStream stream =new FileInputStream(loc);
		Workbook w=new XSSFWorkbook(stream);
		Sheet s=w.getSheet("sheet1");
		Row r =s.getRow(0);
		Cell c=r.getCell(0);
			String s1= c.getStringCellValue();
		if(s1.equals("india")) {
			c.setCellValue("srilatha");
			
		}
		FileOutputStream o= new FileOutputStream(loc);
		w.write(o);
		System.out.println("written");
		// to replace a value 
		
		
		    
		
		
	}


}
