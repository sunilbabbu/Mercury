package com.gs.selenium.testcases;

import org.openqa.selenium.*;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;
import java.net.URL;
import java.util.concurrent.TimeUnit;
import java.net.MalformedURLException;

public class GridEx {
   public WebDriver driver;
   public String URL, Node;
   protected ThreadLocal<RemoteWebDriver> threadDriver = null;
   
   @Parameters("browser")
   @BeforeTest
   public void launchapp(String browser) throws MalformedURLException {
      String URL = "http://newtours.demoaut.com/";
      
      if (browser.equalsIgnoreCase("firefox")) {
         System.out.println(" Executing on FireFox");
         String Node = "http://192.168.1.26:3337/wd/hub";
         DesiredCapabilities cap = DesiredCapabilities.firefox();
         cap.setBrowserName("firefox");
         
         driver = new RemoteWebDriver(new URL(Node), cap);
         // Puts an Implicit wait, Will wait for 10 seconds before throwing exception
         driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
         
         // Launch website
         driver.navigate().to(URL);
         driver.manage().window().maximize();
      } else if (browser.equalsIgnoreCase("chrome")) {
         System.out.println(" Executing on CHROME");
         DesiredCapabilities cap = DesiredCapabilities.chrome();
         cap.setBrowserName("chrome");
         //cap.setPlatform(Platform.LINUX);
         String Node = "http://192.168.1.26:5557/wd/hub";
         driver = new RemoteWebDriver(new URL(Node), cap);
         driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
         
         // Launch website
         driver.navigate().to(URL);
         driver.manage().window().maximize();
      } else if (browser.equalsIgnoreCase("ie")) {
         System.out.println(" Executing on IE");
         DesiredCapabilities cap = DesiredCapabilities.internetExplorer();
         cap.setBrowserName("internet explorer");
         String Node = "http://192.168.1.26:6667/wd/hub";
         driver = new RemoteWebDriver(new URL(Node), cap);
         driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
         
         // Launch website
         driver.navigate().to(URL);
         driver.manage().window().maximize();
      } else {
         throw new IllegalArgumentException("The Browser Type is Undefined");
      }
   }
   
   @Test
   public void calculatepercent() {
      driver.findElement(By.xpath("//*[@name='userName']")).sendKeys("mercury");
      driver.findElement(By.xpath("//*[@name='password']")).sendKeys("mercury");
      driver.findElement(By.xpath("//*[@name='login']")).click();
   }
   
   @AfterTest
   public void closeBrowser() {
     // driver.quit();
   }
}
