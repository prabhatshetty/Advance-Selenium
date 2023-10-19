package com.exceloperations.demo;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class b_ReadAndWrite {
	static String[][] values;

	public static void main(String[] args) throws IOException {
		String excelData_1 = "D:\\Advance Selenium\\Excel_Operations\\data\\Book.xlsx";
		String excelData_2 = "";
		values = readData(excelData_1);
		getValues();
	}

	public static String[][] readData(String path) throws IOException {
		FileInputStream fis = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		int numberOfRows = sheet.getLastRowNum();
		int numberOfColl = sheet.getRow(0).getLastCellNum();
		String[][] data = new String[numberOfRows][numberOfColl];
		for (int i = 0; i < numberOfRows; i++) {
			for (int j = 0; j < numberOfColl; j++) {
				DataFormatter df = new DataFormatter();
				data[i][j] = df.formatCellValue(sheet.getRow(i).getCell(j));
			}
		}
		return data;
	}

	public static void getValues() {
		for (String[] rows : values) {
			for (String col : rows) {
				System.out.print(col + " ");
			}
			System.out.println();
		}
	}
}
