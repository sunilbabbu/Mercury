package com.gs.selenium.testcases;

import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.testng.annotations.AfterTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import com.aventstack.extentreports.Status;

import baseutil.TestBase;
import pagefunction.LoginPage;
import utility.ExcelReader;

public class LoginTest extends TestBase{
	public static final Logger LOG = LogManager.getLogger(LoginTest.class);
	ExcelReader excelReader = new ExcelReader(testDataPath, "LoginData");
	
	@DataProvider(name = "data-provider")
	 public Object[][] dataProviderMethod() throws Exception {
		Object[][] data = excelReader.getDataFromSpreadSheet();
		return data;
	 }
	/**
	 * This method verify Login Test
	 */
	@Test(dataProvider = "data-provider")
	public void login(String userName,String password) {
		test = extent.createTest("Login Test");
		launchBrowser();
		
		new LoginPage().login(userName,password);
		
	}
	
	@AfterTest
	public void tearDown() {
		driver.close();
		LOG.info("Browser closed");
		test.log(Status.INFO, "Browser closed");
		extent.flush();
	}
}
