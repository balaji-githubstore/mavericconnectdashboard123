package com.maveric.mavericconnecttest;

import java.io.IOException;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Optional;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import com.maveirc.utilities.PropertiesHandler;
import com.maveric.pageobjects.LoginPage;
import com.maveric.pageobjects.MavericDashboardPage;
import com.maveric.setup.DriverSetup;

public class AddTimeCartTest extends DriverSetup {
	
	@Test
	public void timeCartTest() throws IOException
	{
		Properties prop= PropertiesHandler.readProperties("utilities/data.properties");
		driver.get(prop.getProperty("url"));
		
		LoginPage login=new LoginPage(driver);
		login.sendUserName(prop.getProperty("username"));
		login.sendPassword(prop.getProperty("password"));
		login.clickOnLogin();
		
		MavericDashboardPage dashboard=new MavericDashboardPage(driver,wait);
		dashboard.mouserHoverAndClickOnAddToTimeCart(); 
		
		//fill form
	boolean check=	driver.getPageSource().contains("QMS");
		
		Assert.assertTrue(check,"Assertion on Add Time Cart Test");
		
	}
}
