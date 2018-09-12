package com.practise.tests;

import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;
import com.practiseAPI.actions.GoogleSheetReadAndWrite;

public class TestGoogleSheetsRead {
	
	GoogleSheetReadAndWrite sheetReadWriteToFile = new GoogleSheetReadAndWrite(); 
	
	@BeforeClass
	public void setUp() {
		sheetReadWriteToFile.retriveDataFromConfig();
	}
	
	@Test
	public void Test_01_testGetAndPrintAllData() {
		sheetReadWriteToFile.getAndPrintAllData();
	}
	
	@Test
	public void Test_02_testWriteTheDataIntoFile() {
		sheetReadWriteToFile.writeTheDataIntoFile();
	}
}
