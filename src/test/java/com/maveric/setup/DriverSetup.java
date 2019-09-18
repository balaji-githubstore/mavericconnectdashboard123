package com.maveric.setup;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Optional;
import org.testng.annotations.Parameters;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;

public class DriverSetup {

	public WebDriver driver;
	public WebDriverWait wait;
	
	// builds a new report using the html template
	public static ExtentHtmlReporter htmlReporter;
	public static ExtentReports extent;
	// helps to generate the logs in test report.
	public static ExtentTest test;

	@BeforeSuite
	@Parameters({ "browser" })
	public void reportSetUp(@Optional("ch") String browser) {
		// initialize the HtmlReporter
		htmlReporter = new ExtentHtmlReporter(System.getProperty("user.dir") + "/extent-reportt/testReport.html");
		// initialize ExtentReports and attach the HtmlReporter
		extent = new ExtentReports();
		extent.attachReporter(htmlReporter);

		// To add system or environment info by using the setSystemInfo method.
		extent.setSystemInfo("Browser", browser);

		// configuration items to change the look and feel
		// add content, manage tests etc
		htmlReporter.config().setDocumentTitle("Extent Report Demo");
		htmlReporter.config().setReportName("Test Report");
		htmlReporter.config().setTheme(Theme.STANDARD);
		htmlReporter.config().setTimeStampFormat("EEEE, MMMM dd, yyyy, hh:mm a '('zzz')'");
	}

	@BeforeTest
	@Parameters({ "browser" })
	public void setUp(@Optional("ch") String browser) {
		System.setProperty("webdriver.chrome.driver", "driver/chromedriver.exe");
		System.setProperty("webdriver.gecko.driver", "driver/geckodriver.exe");
		System.setProperty("webdriver.ie.driver", "driver/IEDriverServer.exe");

		launchBrowser(browser);

		wait = new WebDriverWait(driver, 50);
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(2, TimeUnit.SECONDS);

	}

	public void launchBrowser(String browserName) {
		switch (browserName.toLowerCase()) {
		case "ch":
		case "chrome":
			driver = new ChromeDriver();
			break;
		case "ff":
		case "firefox":
			driver = new FirefoxDriver();
			break;
		default:
			driver = new InternetExplorerDriver();
			break;
		}
	}

	
	// extent report start
	@AfterMethod
	public void getResult(ITestResult result) throws IOException {
		if (result.getStatus() == ITestResult.FAILURE) {
			test.log(Status.FAIL, MarkupHelper.createLabel(result.getName() + " FAILED ", ExtentColor.RED));
			test.fail(result.getThrowable());    
			//screenshot
			String screenShotPath = capture(driver, result.getName());
			test.log(Status.FAIL, "Snapshot below:");
			test.addScreenCaptureFromPath(screenShotPath);
		} else if (result.getStatus() == ITestResult.SUCCESS) {
			test.log(Status.PASS, MarkupHelper.createLabel(result.getName() + " PASSED ", ExtentColor.GREEN));
		} else {
			test.log(Status.SKIP, MarkupHelper.createLabel(result.getName() + " SKIPPED ", ExtentColor.ORANGE));
			test.skip(result.getThrowable());
		}
	}
	
	 public String capture(WebDriver driver,String screenShotName) throws IOException 
	    {
	        TakesScreenshot ts = (TakesScreenshot)driver;
	        File source = ts.getScreenshotAs(OutputType.FILE);
	        
	        
	        String date=new Date().toString().replace(":", "-");
	        String dest = System.getProperty("user.dir") +"\\ErrorScreenshots\\"+screenShotName+" "+date+".png";
	        File destination = new File(dest);
	        FileUtils.copyFile(source, destination);        
	                     
	        return dest;
	    }

	@AfterTest
	public void browserQuit() {
		driver.quit();
	}

	@AfterSuite()
	public void tearDown() {
		// to write or update test information to reporter
		extent.flush();
	}
}
