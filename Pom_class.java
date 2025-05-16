package ksp_admin;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.ParseException;
import java.time.Duration;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Pom_class extends Utilies {

	 WebDriver driver;
	    WebDriverWait wait;

	    // Constructor to initialize WebDriver
	    public Pom_class(WebDriver driver) {
	        this.driver = driver;
	        this.wait = new WebDriverWait(driver, Duration.ofSeconds(10)); // Initialize WebDriverWait
	    }
	    
	    public void Master()
	    {
	    	wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a[contains(text(),'Masters')]"))).click();
	    	sleep(1000);
	    }
	    
	    public void apptype()
	    {
	    	wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a[contains(text(),'Applying Types')]"))).click();
	    	sleep(1000);
	    }
	    
	    public void switc()
	    {
	    switchToNewWindow(driver);
	    sleep(1000);
	    }
	    
	    public void add()
	    {
	    	wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a[contains(text(),'Add')]"))).click();
	    	sleep(1000);
	    }
	    
	    public void adddetail() throws AWTException
	    {
	    	
	    
	    	
		    	  FileInputStream fis = null;
		    	    FileOutputStream fileOut = null;
		    	    Robot r=new Robot();
		    	    Actions act = new Actions(driver);
		    	    JavascriptExecutor jss = (JavascriptExecutor) driver;
		    	    try {
		    			// Reading Excel File
		    			FileInputStream fis1 = new FileInputStream("D://KSP_Admin//KSP_ADMIN_automation.xlsx");//"D:\steno\TestDataAPC.xlsx"
		    			XSSFWorkbook workbook = new XSSFWorkbook(fis1);
		    			Sheet sheet = workbook.getSheetAt(0);

		// Select select = null;
		int rowCount = sheet.getPhysicalNumberOfRows();

		// Loop through rows in the Excel sheet
		// int rowCount = sheet.getPhysicalNumberOfRows();

		
		
		
		for (int i = 1; i <= 1; i++) { // Start from row 1 to skip header
			Row row = sheet.getRow(i);
			
			if (row == null) {
				System.out.println("Skipping empty row: " + i);
				continue;
			}

			if (row != null) {
			
				
				String code = getCellValue(row.	getCell(1));
				String Title = getCellValue(row.getCell(2));
				int Oderindexe = (int) row.getCell(3).getNumericCellValue();
				int status = (int) row.getCell(4).getNumericCellValue();
				
			WebElement CODE = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ApplyingType_Code")));
			CODE.sendKeys(code);
			
			WebElement TITLE = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ApplyingType_Title")));
			TITLE.sendKeys(Title);
				
				
			WebElement oderindex = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ApplyingType_OrderIndex")));	
			act.moveToElement(oderindex).click();
			Select s = new Select(oderindex);
			s.selectByIndex(Oderindexe);
			
			
			WebElement Status = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ApplyingType_StatusCode")));
			Select s1 = new Select(Status);
			s1.selectByIndex(status);
			
			
			
//			WebElement Save = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[contains(text(),'Save')]")));	
//			act.moveToElement(Save).click().perform();
//			sleep(1000);
//			
			WebElement Cancel = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[contains(text(),'Cancel')]")));	
			act.moveToElement(Cancel).click().perform();
				
			
			switchToNewWindow(driver);
			
			
			WebElement search = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("search-bar")));	
			act.moveToElement(search).click().perform();		
			search.sendKeys(code);
			
			r.keyPress(KeyEvent.VK_ENTER);
			r.keyRelease(KeyEvent.VK_ENTER);
			
			sleep(2000);
			WebElement view = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[contains(text(),'visibility')]")));
			act.moveToElement(view).click().perform();

			
			switchToNewWindow(driver);
			sleep(2000);
			switchToNewWindow(driver);
			WebElement viewcode = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='code']")));
			String getviewcode = viewcode.getAttribute("value");	
			System.out.println(getviewcode);
			
			WebElement viewtitle = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='title']")));
			String getviewtitle = viewtitle.getAttribute("value");	
			System.out.println(getviewtitle);
			
			WebElement viewoi = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='orderIndex']")));
			String getviewordrindex = viewoi.getAttribute("value");	
			System.out.println(getviewordrindex);
			
	    	WebElement viewstatus = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='status']")));
			String getviewstatus = viewstatus.getAttribute("value");	
			System.out.println(getviewstatus);
			

	    	WebElement back = wait.until(ExpectedConditions.presenceOfElementLocated(By.linkText("arrow_back")));
	    	act.moveToElement(back).click().perform();
	    	
	    	switchToNewWindow(driver);
	    	
//	    	WebElement search1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("search-bar")));	
//			act.moveToElement(search1).click().perform();	
//			r.keyPress(KeyEvent.VK_BACK_SPACE);
//			sleep(2000);
//			r.keyRelease(KeyEvent.VK_BACK_SPACE);
			   Sheet sheet2 = workbook.getSheetAt(1);
			   Row row1 = sheet2.createRow(sheet2.getPhysicalNumberOfRows()); 
			   
			   row1.createCell(1).setCellValue(getviewcode);
			   row1.createCell(2).setCellValue(getviewtitle);
			   row1.createCell(3).setCellValue(getviewordrindex);
			   row1.createCell(4).setCellValue(getviewstatus);
			   
//			   String codeEdit = getCellValue(row.	getCell(5));
//				String TitleEdit = getCellValue(row.getCell(6));
//				int OderindexEdit = (int) row.getCell(7).getNumericCellValue();
//				int statusEdit = (int) row.getCell(8).getNumericCellValue();
//				
//				switchToNewWindow(driver);
//				sleep(1000);
//				WebElement search1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("search-bar")));
//				act.moveToElement(search).click().perform();		
//				search.sendKeys(code);
//				r.keyPress(KeyEvent.VK_ENTER);
//				r.keyRelease(KeyEvent.VK_ENTER);

				   sleep(3000);
				WebElement edit = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("(//span[contains(text(),'edit')])[1]")));
				act.moveToElement(edit).click().perform();

				switchToNewWindow(driver);
			   
			   sleep(3000);
			   

				fileOut = new FileOutputStream("D://KSP_Admin//KSP_ADMIN_automation.xlsx");
			    workbook.write(fileOut);
			
	    }}}
		    			catch (Exception e) {
							// TODO: handle exception
		    				e.printStackTrace();
						}
		    	    
	    	  }
}
