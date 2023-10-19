package com.exceloperations.demo;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class a_StoringRetrivedDataIn2DArray {
	public static void main(String[] args) throws IOException {
		String[][] data = dataProvider();

		for (int i = 0; i < data.length; i++) {
			for (int j = 0; j < data.length; j++) {
				System.out.println(data[i][j]);
			}
		}
	}

	public static String[][] dataProvider() throws IOException {
		FileInputStream fis = new FileInputStream("D:\\Advance Selenium\\Excel_Operations\\data\\Book.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int numberOfSheets = workbook.getNumberOfSheets();
		System.out.println(numberOfSheets);
		int numberOfRows = 0;
		int numberOfCells = 0;
		String data[][] = null;
	
	}
}
