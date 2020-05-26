package org.d;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class scenario3 {
	public static void main(String[] args) throws IOException {
		File loc =new File ("C:\\Users\\Sundeep\\Documents\\workspace-sts-3.9.11.RELEASE\\DATADRIVEN\\excel\\New Microsoft Excel Worksheet.xlsx");
		FileInputStream stream =new FileInputStream(loc);
		Workbook w= new XSSFWorkbook(stream);
		Sheet s= w.getSheet("Sheet1");
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			Row r=s.getRow(i);
			for (int j = 0;j< r.getPhysicalNumberOfCells(); j++) {
				Cell c =r.getCell(j);
				// 3rd scenario
	  			int type=c.getCellType();
				if (type==1) {
					String output =c.getStringCellValue(); 
					System.out.println(output);
				}else if(type==0){
					if(DateUtil.isCellDateFormatted(c)) {
						Date dat=c.getDateCellValue();
						SimpleDateFormat simple=new SimpleDateFormat("dd-mmm-yyyy");
						String output=simple.format(dat);
						System.out.println(output);
					}
					else {
						double numeric =c.getNumericCellValue();
						long l=(long)numeric;
						String num=String.valueOf(l);
						System.out.println(num);
					}
	}

}
	}
}
}

