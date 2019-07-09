package com.model;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel_xlsxData 
{
	public void readExcel(String filename,String sheetname) throws IOException
	{
		
		FileInputStream fis=new FileInputStream(filename);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheet(sheetname);
		XSSFRow row=sheet.getRow(2);
		XSSFCell cell=row.getCell(2);
		System.out.println("cell value is :" + cell.getStringCellValue());
		System.out.println("excelsheet updated in git hub");
		//no of rows
		int rows=sheet.getLastRowNum();
		System.out.println("The no of rows are : " + rows);
		int rowcount=rows+1;
		System.out.println("The actual no of rows are :" +rowcount);
		//no of columns
		int columns=sheet.getRow(rows).getLastCellNum();
		System.out.println("The no of cloumn are :" + columns);
		int arrayexcel[][]=new int[rowcount][columns];
		for(int i=0;i<rowcount;i++)
		{
			for(int j=0;j<columns;j++)
			{
				System.out.println(sheet.getRow(i).getCell(j));
				
				/*DataFormatter dataformat=new DataFormatter();
				System.out.println(dataformat.formatCellValue(sheet.getRow(i).getCell(j)));*/
			}
		}
		
	}
	public static void main(String[] args) throws IOException
	{
		ReadExcel_xlsxData rsd=new ReadExcel_xlsxData();
		rsd.readExcel("C:\\Users\\user\\workspace1\\HandelExcelSheet\\StudentDetails.xlsx", "sheet1");
	}

}
