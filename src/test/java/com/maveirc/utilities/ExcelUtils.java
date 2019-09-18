package com.maveirc.utilities;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {

	public static Object[][] getSheetIntoObject(String fileDescription,String sheetName) throws IOException
	{
		FileInputStream file=new FileInputStream(fileDescription);

		XSSFWorkbook book=new XSSFWorkbook(file);
	    XSSFSheet sheet=book.getSheet(sheetName);
	    XSSFRow row= sheet.getRow(0);
		
	    int rowCount= sheet.getPhysicalNumberOfRows();
	    System.out.println(rowCount);

	    int cellCount= row.getPhysicalNumberOfCells();
	    System.out.println(cellCount);
	    DataFormatter format =new DataFormatter();
	    
	    Object[][] temp=new Object[rowCount-1][cellCount];
	    
	    for(int i=1;i<rowCount;i++)
	    {
	    	row= sheet.getRow(i);
	    	for(int j=0;j<cellCount;j++)
	    	{
	    		XSSFCell cell= row.getCell(j);
	    	    String cellValue= format.formatCellValue(cell);
	    	    temp[i-1][j]=cellValue;
	    	}
	    }
	    
	    return temp;
	}
	
	
}
