package com.offcn.bean;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

public class Excel_POI {

	
	@Test
	public void testWrite() throws Exception {
		
		HSSFWorkbook workBook = new HSSFWorkbook();
		HSSFSheet sheet = workBook.createSheet("表1");
		HSSFRow row = sheet.createRow(0);
		HSSFCell cell = row.createCell(0);
		cell.setCellValue("这是个表");
		workBook.write(new File("D://upload/1.xls"));
		workBook.close();
		System.out.println("success");
	}
	
	@Test
	public void testRead() throws Exception {
		@SuppressWarnings("resource")
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream("D://upload/1.xls"));
		HSSFSheet sheet = workbook.getSheet("表1");
		HSSFRow row = sheet.getRow(0);
		HSSFCell cell = row.getCell(0);
		System.out.println("==="+cell.getStringCellValue());
	}
	
	@Test
	public void testWriteNew() throws Exception {
		
		HSSFWorkbook workBook = new HSSFWorkbook();
		HSSFSheet sheet = workBook.createSheet("表1");
		HSSFRow row = sheet.createRow(0);
		HSSFCell cell = row.createCell(0);
		cell.setCellValue("这是个新版本表");
		workBook.write(new File("D://upload/1.xlsx"));
		workBook.close();
		System.out.println("success");
	}
	
	
	@Test
	public void testWorkbook() throws Exception {
		Workbook workbook = WorkbookFactory.create(new File("D://upload/1.xls"));
		Sheet sheet = workbook.getSheet("表1");
		Row row = sheet.getRow(0);
		Cell cell = row.getCell(0);
		System.out.println("===="+ cell.getStringCellValue());
	}
	
	
}
