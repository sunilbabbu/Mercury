package utility;

import java.io.File;
import java.io.IOException;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;

import org.apache.commons.io.FileUtils;
import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebElement;

import com.aventstack.extentreports.Status;

import baseutil.TestBase;

public class SeleniumUtility extends TestBase{
	public static final Logger LOG = LogManager.getLogger(SeleniumUtility.class);	
	
	
	public void sendkeys(String xpath,String value,String elementName) {
		WebElement element  =getWebElement(xpath);
		element.sendKeys(value);
		highlightElement(element);
		LOG.info("Value Entered successfully in "+elementName);
		test.log(Status.INFO, "Value Entered successfully in "+elementName);
	}
	public void click(String xpath,String elementName) {
		WebElement element  =getWebElement(xpath);
		highlightElement(element);
		element.click();
		LOG.info("Clicked successfully on "+elementName);
		test.log(Status.INFO, "Clicked successfully on "+elementName);
	}
	public String screenshot(String screenshotName) throws IOException {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
		Timestamp timestamp = new Timestamp(System.currentTimeMillis());
		String date = sdf.format(timestamp);
		String filePath = screenshotPath+screenshotName+date+".png";
		File scrFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
		FileUtils.copyFile(scrFile, new File(filePath));
		return filePath;
	}
	public void highlightElement(WebElement element) {
		((JavascriptExecutor)driver).executeScript("arguments[0].style.border='3px solid green'", element);
	}
	public WebElement getWebElement(String xpath) {
		WebElement element  = driver.findElement(By.xpath(xpath));
		return element;
	}

}
