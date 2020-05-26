package org.d;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Scenario5 {
	public static void main(String[] args) throws IOException  {
		File loc =new File ("C:\\Users\\Sundeep\\Documents\\workspace-sts-3.9.11.RELEASE\\DATADRIVEN\\excel\\New Microsoft Excel Worksheet.xlsx");
		Workbook w=new XSSFWorkbook();
		Sheet s=w.createSheet("Sheet1");
		Row r =s.createRow(1);
		Cell c=r.createCell(1);
		c.setCellValue("india");
		FileOutputStream out=new FileOutputStream(loc);
		w.write(out);
		System.out.println("written");
		/// to write a data in excel 
		
		
		
		
		
		
		
		
		


	}

}



