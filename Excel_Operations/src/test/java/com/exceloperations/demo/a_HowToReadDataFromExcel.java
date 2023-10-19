package com.exceloperations.demo;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class a_HowToReadDataFromExcel {
	@SuppressWarnings("unlikely-arg-type")
	public static void main(String[] args) throws IOException {
		FileInputStream fis = new FileInputStream("D:\\Advance Selenium\\Excel_Operations\\data\\Book.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		/* Getting Number of Sheets */
		int numberOfSheets = workbook.getNumberOfSheets();
		// System.out.println(numberOfSheets);

		/** Iterating through Sheets */
		for (int i = 0; i < numberOfSheets; i++) {
			String sheetName = workbook.getSheetName(i);

			/** Going to the Particular Sheet */
			if (sheetName.equalsIgnoreCase("Sheet1")) {
				XSSFSheet sheet = workbook.getSheetAt(i);

				int numberOfRows = sheet.getLastRowNum();
				System.out.println(numberOfRows);

				int numberOfColumns = sheet.getRow(0).getLastCellNum();
				System.out.println(numberOfColumns);

				/** Iterating through Rows and Column */
				for (int j = 0; j <= numberOfRows; j++) {
					for (int k = 0; k <= numberOfColumns; k++) {
						DataFormatter format = new DataFormatter();
						String val = format.formatCellValue(sheet.getRow(j).getCell(k));
						System.out.print(val + "              ");
					}
					System.out.println("");
				}
			}
		}
	}
}
