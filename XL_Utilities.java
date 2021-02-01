package com.testing.utilities;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XL_Utilities {
	public FileInputStream fi;
	public FileOutputStream fo;
	public XSSFWorkbook workbook;
	public XSSFSheet sheet;
	public  Row row;
	public Cell cell;
	public CellStyle style;
	String path=null;
	public XL_Utilities(String path)
	{
		this.path=path;
	}
	public int getRowCount(String sheetName) throws Exception    
	{
		fi=new FileInputStream(path);
		workbook =new XSSFWorkbook(fi);
		sheet=workbook.getSheet(sheetName);
		int rowcount=sheet.getLastRowNum();
		fi.close();
		workbook.close();
		return rowcount;
	}
	public int getCellCount(String sheetName, int rowrum) throws Exception 
	{
		fi=new FileInputStream(path);
		workbook =new XSSFWorkbook(fi);
		sheet=workbook.getSheet(sheetName);
		row=sheet.getRow(rowrum);
		int cellcount=row.getLastCellNum();
		
		workbook.close();
		fi.close();
		return cellcount;
	}
	public  String  xls_Reader(String sheetName,int rownum,int column) throws Exception
	{
		fi=new FileInputStream(path);
		workbook =new XSSFWorkbook(fi);
		sheet=workbook.getSheet(sheetName);
		row=sheet.getRow(rownum);
		cell=row.getCell(column);
		
		DataFormatter formatter =new DataFormatter();
		
		String data;
		try {
			data=formatter.formatCellValue(cell);
		}
		catch(Exception e)
		{
			data="";
		}
		workbook.close();
		fi.close();
		return data;
	}
	public void setCellData(String sheetName,int rownum,int column,String data) throws Exception
	{
		fi=new FileInputStream(path);
		workbook =new XSSFWorkbook(fi);
		sheet=workbook.getSheet(sheetName);
		row=sheet.getRow(rownum);
		cell=row.createCell(column);
		cell.setCellValue(data);
		
		fo=new FileOutputStream(path);
		workbook.write(fo);
		workbook.close();
		fi.close();
		fo.close();
	}
	public void fillGreenColor(String sheetName,int rownum,int column) throws Exception
	{
		fi=new FileInputStream(path);
		workbook =new XSSFWorkbook(fi);
		sheet=workbook.getSheet(sheetName);
		row=sheet.getRow(rownum);
		cell=row.getCell(column);
		
		style=workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		cell.setCellStyle(style);
		workbook.write(fo);
		workbook.close();
		fi.close();
		fo.close();	
	}
}
