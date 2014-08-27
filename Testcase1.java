package project;

import java.io.File;
import java.io.IOException;
import java.net.URL;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;

import org.openqa.selenium.Platform;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.remote.Augmenter;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.AfterClass;

import corporateadmin.WebDriverMethods;

public class Testcase1 {
	
//	WebDriver  driver;
	String Url = "https://play.google.com/store";
	WebDriver driver = new FirefoxDriver();
	WrapMethods wm = new WrapMethods(driver);
			
  @Test
  public void f() throws BiffException, WriteException, IOException {
	
		int i= 0;
		for (int j=1; j<3;j++){
			
		wm.launchUrl(Url);
		wm.clickValueById("gb_70");
		wm.clickValueById("link-signup");
		wm.setValueById("FirstName", wm.getExcelContent("A",(++i)));
		wm.setValueById("LastName",  wm.getExcelContent("A",(++i)));
		wm.setValueById("GmailAddress",  wm.getExcelContent("A",(++i)));
		wm.setValueById("Passwd",  wm.getExcelContent("A",(++i)));
		wm.setValueById("PasswdAgain",  wm.getExcelContent("A",(++i)));
		wm.clickValueByXpath(".//*[@id='BirthMonth']/div[1]");
		wm.clickValueById(":5");
		wm.setValueById("BirthDay", wm.getExcelContent("A",(++i)));
		wm.setValueById("BirthYear",  wm.getExcelContent("A",(++i)));
		wm.clickValueByXpath(".//*[@id='Gender']/div[1]");
		wm.clickValueById(":d");
		wm.setValueById("RecoveryPhoneNumber", wm.getExcelContent("A",(++i)));
		wm.setValueById("RecoveryEmailAddress",  wm.getExcelContent("A",(++i)));
		wm.clickValueById("TermsOfService");
		wm.clickValueById("submitbutton");
		wm.reportToExcel("TEST CASE ONE ITERATION "+j+" FINISHED", "Report");
		
		}
  }
  
  @BeforeClass
  public void beforeClass() throws BiffException,  IOException, WriteException, InterruptedException {

	  wm.cExcelReport();
		 
	/* DesiredCapabilities capabilities = new DesiredCapabilities();
		capabilities.setPlatform(Platform.WINDOWS);
		capabilities.setBrowserName(DesiredCapabilities.firefox().getBrowserName());
		
		driver = new RemoteWebDriver(new URL("http://192.168.0.5:7777/wd/hub"), capabilities);*/
}

  @AfterClass
  public void afterClass() {
	  wm.finished();
  }

}
