package com.exceloperations.demo;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class b_RealScenario {
	@Test(dataProvider = "dataProvider")
	public void TestCase_1(String usn, String pwd) throws InterruptedException {
		if (usn != null && pwd != null) {
			ChromeOptions option= new ChromeOptions();
			option.addArguments("headless");
			WebDriver driver = new ChromeDriver(option);
			driver.manage().window().maximize();
			driver.get("https://rahulshettyacademy.com/locatorspractice/");
			WebElement userName = driver.findElement(By.id("inputUsername"));
			WebElement passowrd = driver.findElement(By.name("inputPassword"));
			WebElement chkOne = driver.findElement(By.id("chkboxOne"));
			WebElement chkTwo = driver.findElement(By.id("chkboxTwo"));
			WebElement signIn = driver
					.findElement(By.cssSelector("#container > div.form-container.sign-in-container > form > button"));

			Actions a = new Actions(driver);
			a.moveToElement(userName).click().sendKeys(usn).build().perform();
			a.moveToElement(passowrd).click().sendKeys(pwd).build().perform();
			Thread.sleep(4000);
			signIn.click();

			Thread.sleep(4000);
			driver.close();
		}
	}

	@DataProvider
	public String[][] dataProvider() throws IOException {
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
