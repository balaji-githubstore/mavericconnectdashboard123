package com.maveric.pageobjects;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class MavericDashboardPage {
	private By qmsLoc=By.linkText("QMS");
	private By timeCartLoc=By.xpath("//*[text()='Timecard']");
	private By addTimeCartLoc=By.xpath("//*[text()='Add Timecard']");
	
	private WebDriver driver;
	private WebDriverWait wait;
	
	public MavericDashboardPage(WebDriver driver,WebDriverWait wait)
	{
		this.driver=driver;
		this.wait=wait;
	}
	
	public void waitUnitlQMSPresent()
	{
		wait.until(ExpectedConditions.presenceOfElementLocated(qmsLoc));
	}
	
	public String getCurrentTitle()
	{
		return driver.getTitle();
	}
	
	public void mouserHoverAndClickOnAddToTimeCart()
	{
		WebElement timeCartEle= driver.findElement(timeCartLoc);
		WebElement addTimeCartEle= driver.findElement(addTimeCartLoc);
		Actions actions=new Actions(driver);
		actions.moveToElement(timeCartEle).build().perform();
		addTimeCartEle.click();
	}
	

}
