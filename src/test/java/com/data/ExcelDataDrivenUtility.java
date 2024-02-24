package com.data;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataDrivenUtility {

	private String path;

	public ExcelDataDrivenUtility(String path) {

		this.path = path;
	}

	FileInputStream fis;
	XSSFWorkbook workbook;
	Sheet sheet;
	Row row;
	Cell cell;
	FileOutputStream fo;
	XSSFCellStyle style;

	public int getRowCount(String sheetName) throws Exception {
		fis = new FileInputStream(path);
		workbook = new XSSFWorkbook(fis);
		sheet = workbook.getSheet(sheetName);
		int rowCount = sheet.getLastRowNum();
		workbook.close();
		fis.close();

		return rowCount;
	}

	public int getCellCount(String sheetName, int rownum) throws Exception {

		fis = new FileInputStream(path);
		workbook = new XSSFWorkbook(fis);
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(rownum);
		int cellCount = row.getLastCellNum();
		workbook.close();
		fis.close();

		return cellCount;
	}

	public String getCellData(String sheetName, int rownum, int colnum) throws Exception {

		fis = new FileInputStream(path);
		workbook = new XSSFWorkbook(fis);
		sheet = workbook.getSheet(sheetName);

		row = sheet.getRow(rownum);
		cell = row.getCell(colnum);

		DataFormatter formatter = new DataFormatter();
		String data;
		data = formatter.formatCellValue(cell);
		workbook.close();
		fis.close();
		return data;

	}

	public void setCellData(String sheetName, int rownum, int colnum, String data) throws Exception {

		fis = new FileInputStream(path);
		workbook = new XSSFWorkbook(fis);
		sheet = workbook.getSheet(sheetName);

		row = sheet.getRow(rownum);
		cell = row.getCell(colnum);
		cell.setCellValue(data);

		fo = new FileOutputStream(path);
		workbook.close();
		fis.close();
		fo.close();

	}

	public void getCellBackgroundColor(String sheetName, int rownum, int colnum) throws Exception {

		fis = new FileInputStream(path);
		workbook = new XSSFWorkbook(fis);
		sheet = workbook.getSheet(sheetName);

		row = sheet.getRow(rownum);
		cell = row.getCell(colnum);

		style = workbook.createCellStyle();

		style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		cell.setCellStyle(style);
		workbook.close();
		fis.close();
		fo.close();

	}
}