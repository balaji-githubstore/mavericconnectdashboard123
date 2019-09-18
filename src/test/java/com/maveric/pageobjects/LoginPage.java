package com.maveric.pageobjects;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;


public class LoginPage {
	
	private By userLoc= By.cssSelector("#LoginForm_username");
	private By passLoc=By.id("LoginForm_password");
	private By loginLoc=By.xpath("//*[@type='submit']");
	private By errorLoc=By.xpath("//*[text()='Invalid user name']");

	WebDriver driver;
	
	public LoginPage(WebDriver driver)
	{
		this.driver=driver;
	}
	
	public void sendUserName(String userName)
	{
		WebElement userEle= driver.findElement(userLoc);
		userEle.sendKeys(userName);
	}
	
	public void sendPassword(String password)
	{
		WebElement passEle= driver.findElement(passLoc);
		passEle.sendKeys(password);
	}
	
	public void clickOnLogin()
	{
		WebElement loginEle= driver.findElement(loginLoc);
		loginEle.click();
	}
	
	public String getInvalidErrorMessage()
	{
	   WebElement errorEle=driver.findElement(errorLoc);
	   String actualError=errorEle.getText().trim();
	   return actualError;
	}
	
	
}
