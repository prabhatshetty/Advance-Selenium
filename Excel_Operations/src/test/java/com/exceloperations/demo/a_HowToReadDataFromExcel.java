package com.exceloperations.demo;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class a_HowToReadDataFromExcel {
	@SuppressWarnings("unlikely-arg-type")
	public static void main(String[] args) throws IOException {
		FileInputStream fis = new FileInputStream("D:\\Advance Selenium\\Excel_Operations\\data\\Book.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int numberOfSheets = workbook.getNumberOfSheets();
		
		for (int i = 0; i < numberOfSheets; i++) {
			if ((workbook.getSheetName(i)).equals("Sheet1")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				int numberOfRows = sheet.getLastRowNum();
				int numberOfRows1 = sheet.getPhysicalNumberOfRows();
				int numberOfCells = sheet.getRow(0).getLastCellNum();
				int numberOfCells1 = sheet.getRow(0).getPhysicalNumberOfCells();
				System.out.println(numberOfRows);// 5
				System.out.println(numberOfRows1);// 6
				System.out.println(numberOfCells);// 2
				System.out.println(numberOfCells1);// 2
			}
		}
	}
}
