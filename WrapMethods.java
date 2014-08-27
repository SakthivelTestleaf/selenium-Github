package project;


import java.io.File;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementNotVisibleException;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.remote.Augmenter;

public class WrapMethods {
	WebDriver driver;
	int noOfRows;
	String screenshotPath = "E:\\Selenium\\Sel_May\\Program\\reports\\screenshot\\screenshot";
	String reportExcelPath = "E:\\Selenium\\Sel_May\\Program\\reports\\TestReport.xls";
	String testdataExcelPath = "E:\\Selenium\\Sel_May\\Program\\testdata\\TestData.xls";
	
	public WrapMethods (WebDriver driver)
	{
		this.driver = driver;
		
	}
	
	public String getExcelContent(String cl, int i) throws BiffException, IOException, WriteException{
		
			Workbook	wb = Workbook.getWorkbook(new File (testdataExcelPath));
			Sheet	sh = wb.getSheet("Sheet1");
			String r = Integer.toString(i);//converting int to string
			String celllocation = cl.concat(r);//convertinn concating both srting to get he cell location
			String content = sh.getCell(celllocation).getContents();
			
		
		return content;
		
	}
	
	public void cExcelReport() throws BiffException, IOException, WriteException{
		try {
			WritableWorkbook wb = Workbook.createWorkbook(new File(reportExcelPath));
			WritableSheet ws = wb.createSheet("Report", 0);
				ws.setColumnView(0, 10);
				ws.setColumnView(1, 40);
				ws.setColumnView(2, 10);
				ws.setColumnView(3, 14);
			WritableFont cellFont = new WritableFont(WritableFont.COURIER, 13);
			cellFont.setBoldStyle(WritableFont.BOLD);
			WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
			
			cellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
			//cellFormat.setAlignment(Alignment.CENTRE);
			cellFormat.setAlignment(Alignment.JUSTIFY);
			cellFormat.setWrap(true);
			
			Label cl1 = new Label(0, 0, "STEP #",cellFormat);
			Label cl2 = new Label(1, 0, "STEP DESCRIPTION",cellFormat);
			Label cl3 = new Label(2, 0, "STATUS",cellFormat);
			Label cl4 = new Label(3, 0, "SNAPSHOT",cellFormat);
			ws.addCell(cl1);
			ws.addCell(cl2);
			ws.addCell(cl3);
			ws.addCell(cl4);
			wb.write();
			wb.close();
		}catch (Exception e) {
			reportToExcel("Common Exception", "Fail");
		}

	}

	public void reportToExcel(String des, String status ) throws BiffException, IOException, WriteException{
		
		
		try {
			Workbook existingWorkbook = Workbook.getWorkbook(new File(reportExcelPath));
			
			WritableWorkbook workbookCopy = Workbook.createWorkbook(new File(reportExcelPath), existingWorkbook);

/*		openExistingWorkbook(WritableWorkbook workbookCopy);*/
			WritableSheet sheetToEdit = workbookCopy.getSheet("Report");
			WritableCellFormat cellFormat = new WritableCellFormat();
			WritableCellFormat cellFormat1 = new WritableCellFormat();
			WritableCellFormat cellFormat2 = new WritableCellFormat();
			WritableFont cellFont = new WritableFont(WritableFont.ARIAL, 10);
			cellFont.setBoldStyle(WritableFont.BOLD);
			WritableCellFormat cellFormat3 = new WritableCellFormat(cellFont);
			cellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
			cellFormat1.setBorder(Border.ALL, BorderLineStyle.THIN);
			cellFormat2.setBorder(Border.ALL, BorderLineStyle.THIN);
			cellFormat3.setBorder(Border.ALL, BorderLineStyle.THIN);
			
			cellFormat1.setWrap(true);
			cellFormat.setAlignment(Alignment.CENTRE);
			cellFormat1.setAlignment(Alignment.LEFT);
			cellFormat2.setAlignment(Alignment.CENTRE);
			cellFormat3.setAlignment(Alignment.CENTRE);
			
if (status.equals("Pass")){	
	

	

			noOfRows = sheetToEdit.getRows();
			String no = Integer.toString(noOfRows);
			
   //Take SCreen Shot
			File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
			FileUtils.copyFile(src, new File(screenshotPath + noOfRows + ".png"));
  //Add colors
			cellFormat2.setBackground(Colour.GREEN);			
			Label l1 = new Label(0, noOfRows, no,cellFormat);
			Label l2 = new Label(1, noOfRows, des,cellFormat1);
			Label l3 = new Label(2, noOfRows, "Passed", cellFormat2);

			sheetToEdit.addCell(l1);
			sheetToEdit.addCell(l2);
			sheetToEdit.addCell(l3);
			
			sheetToEdit.addCell(new Formula(3, noOfRows,"HYPERLINK(\""+screenshotPath+""+noOfRows+".png\"," + "\"View Snap\")",cellFormat));
								
			workbookCopy.write();
			workbookCopy.close();
			
}
else if(status.equals("Fail")){
	

			noOfRows = sheetToEdit.getRows();
			String no = Integer.toString(noOfRows);
	//Take SCreen Shot
			File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
			FileUtils.copyFile(src, new File(screenshotPath + noOfRows + ".png"));

	//Add colors
			cellFormat2.setBackground(Colour.RED);
			Label l1 = new Label(0, noOfRows, no,cellFormat);
			Label l2 = new Label(1, noOfRows, des,cellFormat1);
			Label l3 = new Label(2, noOfRows, "Failed", cellFormat2);
			sheetToEdit.addCell(l1);
			sheetToEdit.addCell(l2);
			sheetToEdit.addCell(l3);
			sheetToEdit.addCell(new Formula(3, noOfRows,"HYPERLINK(\""+screenshotPath+""+noOfRows+".png\"," + "\"View Snap\")",cellFormat));

			workbookCopy.write();
			workbookCopy.close();
			
}else{
	noOfRows = sheetToEdit.getRows();
	//String no = Integer.toString(noOfRows);
	sheetToEdit.mergeCells(0, noOfRows, 3, noOfRows);
	//Label l1 = new Label(0, noOfRows, no,cellFormat);
	Label l2 = new Label(0, noOfRows, des,cellFormat3);
	//sheetToEdit.addCell(l1);
	sheetToEdit.addCell(l2);
	
	workbookCopy.write();
	workbookCopy.close();
	
}
		}catch (BiffException e) {
			reportToExcel("BiffException", "Fail");
			
		}catch (IOException e) {
			reportToExcel("IOException", "Fail");
			
		}catch (WriteException e) {
			reportToExcel("WriteException", "Fail");
			
		}catch (Exception e) {
			reportToExcel("Common Exception", "Fail");
		}

	
	}
	
	public void takeSnap() throws IOException, BiffException, WriteException {
		//i++;
		
		/*Workbook existingWorkbook = Workbook.getWorkbook(new File("D:\\TestReprot.xls"));
		WritableWorkbook workbookCopy = Workbook.createWorkbook(new File("D:\\TestReprot.xls"), existingWorkbook);
		WritableSheet sheetToEdit = workbookCopy.getSheet("Report");*/
		/*Workbook wb = Workbook.getWorkbook(new File("D:\\TestReprot.xls"));
		Sheet ws = wb.getSheet("Report");
		
		
		int i = ws.getRows();*/
		//WebDriver ad = new Augmenter().augment(driver);
		File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

		// Copy file from memory to physical disk
		FileUtils.copyFile(src, new File("D:\\screenshot" + noOfRows + ".png"));

	}
	
	public void launchUrl (String url) throws IOException, BiffException, WriteException, ElementNotVisibleException
	{
		try {
			driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			driver.manage().window().maximize();
			driver.get(url);
			String desPass = "The url is launched";
			reportToExcel(desPass, "Pass");
		} catch (Exception e) {
			String desFail = "The url is not launched";
			reportToExcel(desFail, "Fail");
		}
	}
	
	public void clickValueById (String Id) throws IOException, BiffException, WriteException
	{
		try {
			
			driver.findElement(By.id(Id)).click();
			String desPass = "The element with the " + Id + " exist and clicked";
			reportToExcel(desPass,"Pass");
		} catch (NoSuchElementException e) {
			String desFail = "The element with "+Id +" is not displayed";
			reportToExcel(desFail, "Fail");
		}catch ( ElementNotVisibleException e) {
			String desFail = "The element with "+Id +" is not displayed";
			reportToExcel(desFail, "Fail");
		}catch (Exception e) {
			reportToExcel("Common Exception", "Fail");
		}
	}
	
	public void setValueById (String Id, String Value) throws IOException, BiffException, WriteException
	{
		try {
			
			driver.findElement(By.id(Id)).sendKeys(Value);
			String desPass =  "The element with the " + Id + " exist and " + Value + " is set";
			reportToExcel(desPass, "Pass");
		} catch (NoSuchElementException e) {
			String desFail = "The element with "+Id +" is not displayed";
			reportToExcel(desFail, "Fail");
		}catch ( ElementNotVisibleException e) {
			String desFail = "The element with "+Id +" is not displayed";
			reportToExcel(desFail, "Fail");
		}catch (Exception e) {
			reportToExcel("Common Exception", "Fail");
		}
	}
	
	public void clickValueByXpath (String Id) throws IOException, BiffException, WriteException
	{
		try {
			
			driver.findElement(By.xpath(Id)).click();
			String desPass = "The element with the " + Id + " exist and clicked";
			reportToExcel(desPass, "Pass");
		} catch (NoSuchElementException e) {
			String desFail = "The element with "+Id +" is not displayed";
			reportToExcel(desFail, "Fail");
			
		}catch ( ElementNotVisibleException e) {
			String desFail = "The element with "+Id +" is not displayed";
			reportToExcel(desFail, "Fail");
		}catch (Exception e) {
			reportToExcel("Common Exception", "Fail");
		}
	}
	
	public void finished(){
		driver.quit();
	}
	


	
	
}
