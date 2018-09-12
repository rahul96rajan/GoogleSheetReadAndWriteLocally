package com.practiseAPI.actions;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.List;

import com.practiseAPI.util.ConfigReaderUtil;
import com.practiseAPI.util.GoogleSheetReaderUtil;
import com.practiseAPI.util.Xlsx_Write_Util;

public class GoogleSheetReadAndWrite {

	String range;
	String sheetID;
	String excelFile;
	String excelSheetName;
	List<List<Object>> values = null;

	public void retriveDataFromConfig() {
		ConfigReaderUtil.initialiseProperties();
		range = ConfigReaderUtil.getProperty("range");
		sheetID = ConfigReaderUtil.getProperty("spreadsheetId");
		excelFile = ConfigReaderUtil.getProperty("excel_loc");
		excelSheetName = ConfigReaderUtil.getProperty("sheet_name");
	}

	public void getAndPrintAllData() {
		retriveDataFromConfig();
		try {
			values = GoogleSheetReaderUtil.retriveAllValues(range, sheetID);
		} catch (GeneralSecurityException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		System.out.println("\n################--THE VALUE FROM THE SHEET--################\n");
		for (List row : values) {
			for (int j = 0; j < row.size(); j++) {
				System.out.print(row.get(j) + ", ");
			}
			System.out.println();
		}
	}
	
	public void writeTheDataIntoFile() {
		Xlsx_Write_Util.init(excelSheetName);
		Xlsx_Write_Util.writeDataToExcelFile(values, excelFile);
		System.out.println("\nThese are Written to File at location : \n" + excelFile + "\n\n");
	}
}
