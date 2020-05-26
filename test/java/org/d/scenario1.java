package org.d;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class scenario1 {
	public static void main(String[] args) throws IOException {
		File loc =new File ("C:\\Users\\Sundeep\\Documents\\workspace-sts-3.9.11.RELEASE\\DATADRIVEN\\excel\\New Microsoft Excel Worksheet.xlsx");
		FileInputStream stream =new FileInputStream(loc);
		Workbook w= new XSSFWorkbook(stream);
		Sheet s= w.getSheet("Sheet1");
		Row r=s.getRow(1);
		Cell c=r.getCell(0);
		System.out.println(c);
		
		
		
		
		
		
		
	}

}
