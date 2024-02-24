package com.data;

import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Realtime {

	@Test(dataProvider = "LoginData")
	public void loginTest(String user, String pass, String value) {
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		driver.manage().window().maximize();

		driver.get("https://www.nopcommerce.com/en/login?returnUrl=%2Fen%2Fdemo");
		driver.findElement(By.id("Username")).sendKeys(user);
		driver.findElement(By.id("Password")).sendKeys(pass);
		driver.findElement(By.cssSelector(".btn.blue-button")).click();

	}

	@DataProvider(name = "LoginData")
	public String[][] getData() throws Exception {

		/*
		 * String loginData[][] = { { "admin@yourstore.com", "admin", "valid" }, {
		 * "adm@yourstore.com", "admin", "invalid" }, { "ad@yourstore.com", "admin",
		 * "invalid" }
		 * 
		 * };
		 */

		String path = "C:\\Users\\katak\\eclipse-workspace\\PavanDatadriven\\orgdata.xlsx";

		ExcelDataDrivenUtility ex = new ExcelDataDrivenUtility(path);
		int totalrows = ex.getRowCount("Sheet1");
		int totalcols = ex.getCellCount("sheet1", 1);
		String loginData[][] = new String[totalrows][totalcols];

		for (int i = 1; i <= totalrows; i++) {
			for (int j = 0; j < totalcols; j++) {

				loginData[i - 1][j] = ex.getCellData("sheet1", i, j);

			}
		}

		return loginData;
	}
}
