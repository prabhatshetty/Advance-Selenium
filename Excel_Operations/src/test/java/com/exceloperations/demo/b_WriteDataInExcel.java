package com.exceloperations.demo;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class b_WriteDataInExcel {
	public static void main(String[] args) throws IOException {
		String path = "D:\\Advance Selenium\\Excel_Operations\\data\\Book6.xlsx";
		FileOutputStream fos = new FileOutputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook();
		// XSSFSheet sheet = workbook.createSheet("Sheet_2");
		XSSFSheet sheet = workbook.getSheet("Sheet_1");

		Object[][] data = { { "Test_1", 123 }, { "Test_2", 456 } };
		int numberOfRows = data.length;
		int numberOfCells = data[0].length;
		System.out.println(numberOfRows);

//		XSSFRow row = sheet.createRow(numberOfRows);
//		XSSFCell cell = row.createCell(numberOfCells);

		for (int i = 0; i < numberOfRows; i++) {
			XSSFRow row = sheet.createRow(i);
			for (int j = 0; j < numberOfCells; j++) {
				XSSFCell cell = row.createCell(j);
				// Object val =data[i][j];
				for (Object[] row_1 : data) {
					for (Object cell_1 : row_1) {
						cell.setCellValue(cell_1.toString());
					}
				}
			}
		}
		workbook.write(fos);
		fos.close();
		System.out.println("kodijwfuhgy");
	}
}
