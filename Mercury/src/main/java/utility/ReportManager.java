package utility;

import java.sql.Timestamp;
import java.text.SimpleDateFormat;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;

public class ReportManager {
	
	public static ExtentReports getReportInstance() {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
		Timestamp timestamp = new Timestamp(System.currentTimeMillis());
		String date = sdf.format(timestamp);
		ExtentHtmlReporter htmlReporter = new ExtentHtmlReporter(System.getProperty("user.dir")+"\\Reports\\SeleniumReport"+date+".html");
		ExtentReports extent = new ExtentReports();
		extent.attachReporter(htmlReporter);
		return extent;
	}

}
