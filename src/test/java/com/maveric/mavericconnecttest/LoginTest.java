package com.maveric.mavericconnecttest;
import java.io.IOException;
import java.util.Properties;
import java.util.concurrent.TimeUnit;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.Reporter;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Optional;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import com.maveirc.utilities.ExcelUtils;
import com.maveirc.utilities.PropertiesHandler;
import com.maveric.pageobjects.LoginPage;
import com.maveric.pageobjects.MavericDashboardPage;
import com.maveric.setup.DriverSetup;

public class LoginTest extends DriverSetup {
	
	/**
	 * Invalid Credentials
	 * throws IOException
	 * **/
	@Test//(groups = {"low"},dependsOnMethods = "validCredentials")
	public void invalidCredentialTest() throws IOException
	{
		try
		{
			test=extent.createTest("invalidCredentialTest");
			
			Properties prop= PropertiesHandler.readProperties("utilities/data.properties");
			
			String url = prop.getProperty("url");
			String userName = prop.getProperty("invalidusername");
			String password = prop.getProperty("invalidpassword");
				String expectedError=prop.getProperty("expectederror");
			driver.get(url);
			
			LoginPage login=new LoginPage(driver);
			
			login.sendUserName(userName);
			login.sendPassword(password);
			login.clickOnLogin();
			
			String actualError= login.getInvalidErrorMessage();
			Assert.assertEquals(actualError, expectedError,"Assertion on InvalidCredential");
			
		}
		catch (Exception e) {
			Reporter.log("Exception on Invalid Credentials "+e.getMessage());
			Assert.fail();
		}
		
		
	}
	
	@Test(priority = 2,timeOut = 10000,invocationCount = 5,invocationTimeOut = 50000,enabled = false,groups = {"high","credential"})
	public void validCredentialTest () throws IOException {
		
		test=extent.createTest("validCredentialTest");
		Properties prop= PropertiesHandler.readProperties("utilities/data.properties");
		
		String url = prop.getProperty("url");
		String userName = prop.getProperty("username");
		String password = prop.getProperty("password");
		String expectedTitle=prop.getProperty("expectedtitle");		
		
		driver.get(url);
		
		LoginPage login=new LoginPage(driver);
		
		login.sendUserName(userName);
		
		login.sendPassword(password);
		
		login.clickOnLogin();
		
		MavericDashboardPage dashboard=new MavericDashboardPage(driver, wait);
		dashboard.waitUnitlQMSPresent();
		
		String actualTitle=dashboard.getCurrentTitle();
		System.out.println(actualTitle);
		
		Assert.assertEquals(actualTitle, expectedTitle,"Assertion on ValidCredentialTest");
		
	}
	
//	@DataProvider(name = "InvalidCredentials")
//	public Object[][] getTestDataForInvalidCredentials(){
//		Object[][] temp=new Object[3][2];
//		//i - no of times need to run
//		//j -- no of parameters
//		
//		temp[0][0]="balaji0017@gmail.com";
//		temp[0][1]="password123";
//		
//		temp[1][0]="bala@gmail.com";
//		temp[1][1]="password456";
//		
//		temp[2][0]="bala@gmail.com";
//		temp[2][1]="password456";
//		
//		return temp;
//	}
	
	
	@DataProvider(name = "InvalidCredentials")
	public Object[][] getTestDataForInvalidCredentials() throws IOException{
		Object[][] temp = ExcelUtils.getSheetIntoObject("utilities/MavericDashboardTest.xlsx", "InvalidCredentials");
		return temp;
	}
	
	
	@Test(dataProvider = "InvalidCredentials",priority = 1)
	public void invalidCredentialsTest(String userName,String password) throws IOException
	{
		try
		{
			test=extent.createTest("invalidCredentialsTest"+userName);
			Properties prop= PropertiesHandler.readProperties("utilities/data.properties");
			
			String url = prop.getProperty("url");
			
			String expectedError=prop.getProperty("expectederror");
			driver.get(url);
			
			LoginPage login=new LoginPage(driver);
			
			login.sendUserName(userName);
			login.sendPassword(password);
			login.clickOnLogin();
			
			String actualError= login.getInvalidErrorMessage();
			Assert.assertEquals(actualError, expectedError,"Assertion on InvalidCredential");
			
		}
		catch (Exception e) {
			Reporter.log("Exception on Invalid Credentials "+e.getMessage());
			Assert.fail();
		}
		
		
	}
	
}




