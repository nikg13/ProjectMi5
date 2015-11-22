
import java.io.*;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;


import java.text.*;


public class ReadExcel {

	public static void main(String[] args) throws Exception	{
		// TODO Auto-generated method stub 
		//hi nik
		File excel_file  = new File ("C:\\Users\\nik\\workspace\\HelloWorld\\TestIP.xlsx");
		FileInputStream FIS = new FileInputStream(excel_file);
		
		XSSFWorkbook wb = new XSSFWorkbook(FIS);
		XSSFSheet ws = wb.getSheet("Sheet1");
		
		int rowNum = ws.getLastRowNum() + 1;
		int colNum = ws.getRow(0).getLastCellNum();
		
		String[][] data = new String [rowNum][colNum];
		
		System.out.println("R: "+rowNum + "C: " +colNum);
		
		for (int i = 0 ; i < rowNum ; i++)
		{
			XSSFRow row = ws.getRow(i);
		
			for (int j = 0 ; j < colNum ; j++ )
			{
				XSSFCell cell = row.getCell(j);
				String value = celltostring(cell);
				data[i][j] = value ; 
				System.out.println(data[i][j]);
			}
		}
		
		writeExcel(data,rowNum,colNum);
		
		
	}

	public static String celltostring(XSSFCell cell)
	{
		int type ;
		Object result = null;
		type = cell.getCellType();
		
		switch (type)
		{
		case 0 : 
			result = cell.getNumericCellValue();
		case 1 :
			 result = cell.getStringCellValue();
		}
		return result.toString();
		
	}
	
	public static void writeExcel(String data[][],int rowNum,int colNum) throws IOException
	{
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("sheetOne");
				
				for (int i = 0 ; i < rowNum ; i++)
				{
					HSSFRow row = sheet.createRow(i);
					for (int j = 0 ; j < colNum ; j++ )
					{
						HSSFCell cell = row.createCell(j);
						cell.setCellValue(data[i][j]);
					}
						
				}
					
		workbook.write(new FileOutputStream("Test_op.xls"));
		workbook.close();
		
	}
	
	
}
