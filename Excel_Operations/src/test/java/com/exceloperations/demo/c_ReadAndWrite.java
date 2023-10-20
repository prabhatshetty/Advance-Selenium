package com.exceloperations.demo;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class c_ReadAndWrite {
	static String[][] values;

	public static void main(String[] args) throws IOException {
		String excelData_1 = "D:\\Advance Selenium\\Excel_Operations\\data\\Book.xlsx";
		String excelData_2 = "D:\\Advance Selenium\\Excel_Operations\\data\\demo.xlsx";
		values = readData(excelData_1);
		getValues(values);
		writeData(excelData_2);
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

	public static void getValues(String[][] arr) {
		for (String[] rows : arr) {
			for (String col : rows) {
				System.out.print(col + " ");
			}
			System.out.println();
		}
	}

	public static void writeData(String path) throws IOException {
		FileOutputStream fos = new FileOutputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Sheet_2");
		fos.flush();
		fos.close();
		
	}
}
