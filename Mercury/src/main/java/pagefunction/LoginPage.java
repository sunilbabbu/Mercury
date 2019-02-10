package pagefunction;

import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;

import com.aventstack.extentreports.MediaEntityBuilder;
import com.aventstack.extentreports.Status;

import baseutil.TestBase;
import pageobjects.Login;
import utility.SeleniumUtility;

public class LoginPage extends TestBase{
public static final Logger LOG = LogManager.getLogger(LoginPage.class);
SeleniumUtility selUtil = new SeleniumUtility();	


public void login(String userName,String password) {
		try {
			selUtil.sendkeys(Login.USERNAME, userName,"UserName textbox");
			selUtil.sendkeys(Login.PASSWORD, password,"Password textbox");
			test.log(Status.INFO, "Login Credentials Entered",MediaEntityBuilder.createScreenCaptureFromPath(selUtil.screenshot("Login")).build());
			selUtil.click(Login.LOGINBUTTON, "Login Button");
			
		}
		catch(Exception e) {
			LOG.error("Error in Login Page");
			test.log(Status.FAIL, "Error in Login Page");
		}
	}
}
