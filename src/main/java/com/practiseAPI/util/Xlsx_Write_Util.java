package com.practiseAPI.util;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Xlsx_Write_Util {
	static Workbook workbook = new XSSFWorkbook();
	CreationHelper createHelper = workbook.getCreationHelper();
	static Sheet sheet;
	static Font headerFont = workbook.createFont();
	static Row headerRow;
	static CellStyle headerCellStyle;

	public static void init(String nameOfSheetToBeMade) {
		sheet = workbook.createSheet(nameOfSheetToBeMade);
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.BLACK.getIndex());
		headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);
		headerRow = sheet.createRow(0);
	}

	public static void writeDataToExcelFile(List<List<Object>> values, String excelFilePath) {
		List<Object> column = values.get(0);

		int noOfColumn = values.get(0).size();
		for (int i = 0; i < noOfColumn; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue((String) column.get(i));
			cell.setCellStyle(headerCellStyle);
		}

		int rowNum = 1;
		for (List<Object> val : values) {
			Row row = sheet.createRow(rowNum++);
			for (int i = 0; i < column.size(); i++) {
				row.createCell(i).setCellValue((String) val.get(i));
			}
		}

		for (int i = 0; i < column.size(); i++) {
			sheet.autoSizeColumn(i);
		}

		FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream(excelFilePath);
			workbook.write(fileOut);
			fileOut.close();
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
