package com.exceloperations.demo;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class a_StoringRetrivedDataIn2DArray {
	public static void main(String[] args) throws IOException {
		String[][] data = dataProvider();

//		for (String[] row : data) {
//			for (String col : row) {
//				System.out.print(col+" ");
//			}
//			System.out.println("");
//		}
		for (int i = 1; i < data.length; i++) {
			for (int j = 1; j < data[i].length; j++) {
				System.out.println(data[i][j]);
			}
		}
	}

	public static String[][] dataProvider() throws IOException {
		FileInputStream fis = new FileInputStream("D:\\Advance Selenium\\Excel_Operations\\data\\Book.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		XSSFSheet sheet = workbook.getSheet("Sheet1");

		int numberOfRows = sheet.getLastRowNum();
		int numberOfColumns = sheet.getRow(0).getLastCellNum();
		String[][] data = new String[numberOfRows][numberOfColumns];
		for (int i = 1; i < numberOfRows; i++) {
			for (int j = 0; j < numberOfColumns; j++) {
				DataFormatter df = new DataFormatter();
				data[i][j] = df.formatCellValue(sheet.getRow(i).getCell(j));
				System.out.println(data[i][j]);
			}
		}
		return data;
	}
}
