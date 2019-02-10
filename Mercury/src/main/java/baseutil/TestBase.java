package baseutil;

import java.util.concurrent.TimeUnit;

import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.log4j.xml.DOMConfigurator;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;

import utility.ReportManager;

public class TestBase {
	
	/** WebDriver instance */
	public static WebDriver driver =null;
	public static final Logger LOG = LogManager.getLogger(TestBase.class);
	public String testDataPath = System.getProperty("user.dir")+"\\src\\test\\resources\\MercuryTestData.xlsx";
	public String screenshotPath = System.getProperty("user.dir")+"\\Reports\\screenshots\\";
	public ExtentReports extent  = ReportManager.getReportInstance();
	public static ExtentTest test=null;
	static{
		DOMConfigurator.configure(System.getProperty("user.dir")+"\\src\\test\\resources\\log4j1.xml");
	}
	/**
	 * This method will launch browser.
	 */
	public void launchBrowser() {
		System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir")+"\\src\\test\\resources\\browserdrivers\\chromedriver.exe");
		driver = new ChromeDriver();
		LOG.info("Browser instance created");
		driver.manage().window().maximize();
		LOG.info("Browser Maximized");
		driver.get("http://newtours.demoaut.com/");
		LOG.info("Application Launched");
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	}
}
