package project;

import java.io.IOException;
import java.net.URL;

import jxl.read.biff.BiffException;
import jxl.write.WriteException;

import org.openqa.selenium.Platform;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.remote.Augmenter;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.testng.annotations.Test;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.AfterClass;

public class Testcase2 {
	
	String Url = "https://play.google.com/store";

	WebDriver  driver;


//	 WebDriver driver = new FirefoxDriver();
	WrapMethods wm;
	//WrapMethods wm = new WrapMethods(driver);
	
  @Test
  public void f() throws BiffException, WriteException, IOException {
	  
	  	int ii=0;
		for (int i = 1; i < 3; i++) {
			
			wm.launchUrl(Url);
			wm.clickValueById("gb_70");
			wm.setValueById("Email", wm.getExcelContent("B", ++ii));
			wm.setValueById("Passwd", wm.getExcelContent("B", ++ii));
			wm.clickValueById("signIn");
			wm.clickValueByXpath("//*[@id='gb']/div[1]/div[1]/div[2]/div[5]/div[1]/a/span");
			wm.clickValueById("gb_71");
			wm.reportToExcel("TEST CASE TWO ITERATION "+i+" FINISHED", "Report");
		}
  }
  @BeforeClass
  public void beforeClass() throws BiffException, WriteException, IOException, InterruptedException {
	  System.setProperty("webdriver.chrome.driver", "E:\\Selenium\\Sel_May\\Program\\drivers\\chromedriver.exe");
	  driver = new ChromeDriver();
	  wm = new WrapMethods(driver);
	  
	 // wm.cExcelReport();
	/*  	 DesiredCapabilities capabilities = new DesiredCapabilities();
		capabilities.setPlatform(Platform.WINDOWS);
		capabilities.setBrowserName(DesiredCapabilities.firefox().getBrowserName());
		
		driver = new RemoteWebDriver(new URL("http://192.168.0.5:7777/wd/hub"), capabilities);*/

  }

  @AfterClass
  public void afterClass()  {
	  wm.finished();
	  //driver.quit();
  }

}
