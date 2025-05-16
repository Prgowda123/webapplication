package Steno;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Reporter;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class Home {
	WebDriver driver;
	WebDriverWait wait;
	
	@BeforeClass
	public void OpenBrowser() {
		driver= new ChromeDriver();
		driver.manage().window().maximize();
		wait= new WebDriverWait(driver, Duration.ofSeconds(10));
		
	}
	
	@AfterClass
	public void CloseBrowser() {
		driver.close();
		driver.quit();
	}
	
	@Test(priority=1)
	public void StenoHome() throws InterruptedException, IOException {
		wait= new WebDriverWait(driver, Duration.ofSeconds(10));
		
		driver.get("http://172.10.1.159:9013/");
		Set<String> All_Windows = driver.getWindowHandles();
		int count = All_Windows.size();
		System.out.println(count);
		
		
		for (String text : All_Windows) {
			System.out.println(text);
			driver.switchTo().window(text);
			System.out.println(driver.getTitle());
		}
		
		List<WebElement> All_Links = driver.findElements(By.tagName("a"));
		int count_links = All_Links.size();
		System.out.println(" ");
		System.out.println(count_links);
		
		for (WebElement webElement : All_Links) {
			String url = webElement.getAttribute("href");
			 if (url != null && !url.isEmpty()) {  
	                System.out.println(url);
	            }
			 
			String text_link = webElement.getText();
			System.out.println(text_link);	
		}
	
			WebElement Home = driver.findElement(By.xpath("//a[contains(text(),'Home')]"));
			String home_text = Home.getText();
		
			
			WebElement Broucher = driver.findElement(By.xpath("//a[contains(text(),'Broucher')]"));
			String broucher_text = Broucher.getText();
		
			
			WebElement Notification = driver.findElement(By.xpath("//a[text()='Notification']"));
			String notification_text = Notification.getText();
		
			
			WebElement NewApplication = driver.findElement(By.xpath("//a[@class='nav-link newapplication']"));
			String newApplication_text = NewApplication.getText();
	
			NewApplication.click();
			
			         ArrayList<String> tab= new ArrayList<String>(driver.getWindowHandles());
			     	 driver.switchTo().window(tab.get(1));
			           Thread.sleep(2000);
			
			         driver.switchTo().window(tab.get(1));
			           Thread.sleep(2000);
			         driver.findElement(By.xpath("//button[@class='my-app-btn']")).click();
			           Thread.sleep(2000);
			           
			 WebElement MyApplication = driver.findElement(By.xpath("//a[@class='my-app-btn applicationlogin']"));
			 String myApllication_text = MyApplication.getText();
			
			 MyApplication.click();
			   		
			        ArrayList<String> tab1= new ArrayList<String>(driver.getWindowHandles());
			   		driver.switchTo().window(tab1.get(1));
			   		   Thread.sleep(2000);
			   		driver.findElement(By.xpath("//button[@class='my-app-btn']")).click();
			   	    driver.switchTo().window(tab.get(0));
			   		   Thread.sleep(2000);
			 
			
			WebElement Download_Notification = driver.findElement(By.xpath("//a[@class='btn apply-btn me-2']"));
			String downloadNotification_text = Download_Notification.getText();
			
			Thread.sleep(2000);
			
			WebElement ApplyNow = driver.findElement(By.xpath("//a[@class='btn apply-btn newapplication ms-2']"));
			String applyNow_text = ApplyNow.getText();
		
			ApplyNow.click();
			Thread.sleep(2000);
			
			       ArrayList<String> tab2= new ArrayList<String>(driver.getWindowHandles());
			       driver.switchTo().window(tab2.get(1));
			       driver.findElement(By.xpath("//button[@class='my-app-btn']")).click();
			         Thread.sleep(1000);
			         
			         driver.switchTo().window(tab.get(0));
			      JavascriptExecutor js= (JavascriptExecutor) driver;
			      js.executeScript("window.scrollBy(0,600)");
			        Thread.sleep(2000);
			        
			WebElement Noifi_number_Eng = driver.findElement(By.xpath("//p[contains(text(),'Notification No : 01 / Recruitment-6 / 2024-25')] "));
			String notifi_eng = Noifi_number_Eng.getText();
			
			
			WebElement Noifi_number_kan = driver.findElement(By.xpath("//span[contains(text(),\"C¢ü¸ÀÆZÀ£É ¸ÀASÉå : 01 / £ÉÃªÀÄPÁw-6 / 2024-25\")] "));
			String notifi_kan = Noifi_number_kan.getText();
			
			String Notifi_headline_eng = driver.findElement(By.xpath("//h3[@class='mb-3']")).getText();
		
			
			String Notifi_headline_kan = driver.findElement(By.xpath("//h3[@class='kannada-font mb-4']")).getText();
			
			WebElement Age_Criteria = driver.findElement(By.xpath("//h5[contains(text(),'AGE CRITERIA')]"));
			String ageCriteria_text = Age_Criteria.getText();
		
			String Category = driver.findElement(By.xpath("//th[contains(text(),'Category')]")).getText();
	
			
			String MinimumAge = driver.findElement(By.xpath("//th[contains(text(),'Minimum Age')]")).getText();
			
			String Age01 = driver.findElement(By.xpath("(//th[text()='Age'])[1]")).getText();
			
			String Date01 = driver.findElement(By.xpath("(//th[text()='Date'])[1]")).getText();
			
			String MaximumAge = driver.findElement(By.xpath("//th[contains(text(),'Maximum Age')]")).getText();
			
			String Age02 = driver.findElement(By.xpath("(//th[text()='Age'])[2]")).getText();
			
			String Date02 = driver.findElement(By.xpath("(//th[text()='Date'])[2]")).getText();
			
			String OverAge = driver.findElement(By.xpath("//th[contains(text(),'Overage')]")).getText();
			
			String UnderAge = driver.findElement(By.xpath("//th[contains(text(),'Underage')]")).getText();
			
			WebElement AgeCalculator = driver.findElement(By.xpath("//button[@class='btn age-btn']"));
			String age_calci_text = AgeCalculator.getText();
			
			AgeCalculator.click();
			
			Thread.sleep(2000);
			
			driver.findElement(By.xpath("//input[@placeholder='Select Date']")).click();
			Thread.sleep(2000);
			
			WebElement yearKey = driver.findElement(By.xpath("//span[@class='arrowDown']"));
			    for (int count1=1;count1<=12;count1++) {
				     yearKey.click();
			    }
			
			WebElement month = driver.findElement(By.xpath("//select[@class='flatpickr-monthDropdown-months']"));
			Select select= new Select(month);
			select.selectByVisibleText("September");
			Thread.sleep(2000);
			
			WebElement day = driver.findElement(By.xpath("(//span[@class='flatpickr-day'])[13]"));
			day.click();
			Thread.sleep(1000);
			
			WebElement calculate_age = driver.findElement(By.xpath("//button[@class='btn calculate-age']"));
			calculate_age.click();
			
			WebElement age_result = driver.findElement(By.id("ageResult"));
			String age_result_text = age_result.getText();
			System.out.println(age_result_text);
		
			Thread.sleep(2000);
			
			driver.findElement(By.xpath("//button[@class='btn btn-secondary']")).click();
			Thread.sleep(2000);
			
			js.executeScript("window.scrollBy(0,600)");
			
			String Key_dates = driver.findElement(By.xpath("//h5[contains(text(),'KEY DATES')]")).getText();

			
			String Note = driver.findElement(By.xpath("//h5[contains(text(),'NOTES')]")).getText();
		
			String Gen_info = driver.findElement(By.xpath("//h3[contains(text(),'General Information')]")).getText();
		
			
			String QuickLinks = driver.findElement(By.xpath("//h3[contains(text(),'Quick Links')]")).getText();
		
			String ContactUs = driver.findElement(By.xpath("//h3[contains(text(),'Contact Us')]")).getText();
			
			

			
		Reporter.log("Homepage Titles are completed");
		Thread.sleep(2000);

	    FileInputStream fis1 = new FileInputStream("D://Automation_data//TestData (2).xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis1);
		 Sheet sheet = workbook.getSheetAt(2);
		 int rowNumber = 15; 
		 Row row1 = sheet.createRow(rowNumber);
	  // Row row1 = sheet.createRow(sheet.getPhysicalNumberOfRows()); 
		   
		        row1.createCell(1).setCellValue(home_text);
	            row1.createCell(2).setCellValue(broucher_text);
	            row1.createCell(3).setCellValue(notification_text);
	            row1.createCell(4).setCellValue(newApplication_text);
	            row1.createCell(5).setCellValue(myApllication_text);
	            row1.createCell(6).setCellValue(downloadNotification_text);
	            row1.createCell(7).setCellValue(applyNow_text);
	            row1.createCell(8).setCellValue(notifi_eng);
	            row1.createCell(9).setCellValue(notifi_kan);
	            row1.createCell(10).setCellValue(Notifi_headline_eng);
	            row1.createCell(11).setCellValue(Notifi_headline_kan);
	            row1.createCell(12).setCellValue(ageCriteria_text);
	            row1.createCell(13).setCellValue(Category);
	            row1.createCell(14).setCellValue(MinimumAge);
	            row1.createCell(15).setCellValue(Age01);
	            row1.createCell(16).setCellValue(Date01);
                row1.createCell(17).setCellValue(MaximumAge);
                row1.createCell(18).setCellValue(Age02);
                row1.createCell(19).setCellValue(Date02);
                row1.createCell(20).setCellValue(OverAge);
                row1.createCell(21).setCellValue(UnderAge);
                row1.createCell(22).setCellValue(age_calci_text);
                row1.createCell(23).setCellValue(age_result_text);
                row1.createCell(24).setCellValue(Key_dates);
                row1.createCell(25).setCellValue(Note);
                row1.createCell(26).setCellValue(Gen_info);
                row1.createCell(27).setCellValue(QuickLinks);
                row1.createCell(28).setCellValue(ContactUs);
              
	            
	            FileOutputStream fileOut = new FileOutputStream("D://Automation_data//TestData (2).xlsx");
	            workbook.write(fileOut);
	            fileOut.close();

	}}
