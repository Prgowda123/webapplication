package university;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class DBEID {
	
@Test
public void sample() throws IOException, InterruptedException
{
	ChromeDriver driver = new ChromeDriver();
	
	driver.manage().window().maximize();
	driver.get("https://deb.ugc.ac.in/College/StudentDEBId/Index");

	 FileInputStream fis = null;
    FileOutputStream fileOut = null;

	FileInputStream fis1 = new FileInputStream("D://Automation_data//SVU_Admission_20250401_4.xlsx");//"D:\steno\TestDataAPC.xlsx"
	XSSFWorkbook workbook = new XSSFWorkbook(fis1);
	Sheet sheet = workbook.getSheetAt(0);

	// Select select = null;
	int rowCount = sheet.getPhysicalNumberOfRows();

	// Loop through rows in the Excel sheet
	// int rowCount = sheet.getPhysicalNumberOfRows();
	WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
	JavascriptExecutor jss = (JavascriptExecutor) driver;
	WebElement Applicant_FullName = wait.until(ExpectedConditions.presenceOfElementLocated(By.name("InstituteID")));
	Applicant_FullName.sendKeys("HEI-U-0037");
	
	WebElement password = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("txtPassword")));
	password.sendKeys("INF@123");
	Thread.sleep(6000);
	
	WebElement signin = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("btnLogin")));
	signin.click();
	Thread.sleep(600);
	jss.executeScript("window.scrollBy(0,100)", "");
	
	WebElement button = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a//div[@class='homePage_iconPart text-center']")));
	button.click();
	 

	 
    Set<String> windowHandles = driver.getWindowHandles(); // Now, no NullPointerException
    List<String> tabs = new ArrayList<>(windowHandles);
    driver.switchTo().window(tabs.get(1));
   
	
	Thread.sleep(1000);
	WebElement prioriti = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Year")));
	Thread.sleep(600);
	Select p=new Select(prioriti);
	p.selectByIndex(2);


	WebElement ApplyingType = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("session")));
	Select s=new Select(ApplyingType);
	s.selectByIndex(1);
	
	Thread.sleep(3000);
	try {
	for (int i = 1; i <=1107; i++) { // Start from row 1 to skip header
		Row row = sheet.getRow(i);
		
		if (row == null) {
			System.out.println("Skipping empty row: " + i);
			continue;
		}
        
		if (row != null) {
			
			
			String DEB_id = getCellValue(row.getCell(0));
			
			Sheet sheet2 = workbook.getSheetAt(1);
			   Row row1 = sheet2.createRow(sheet2.getPhysicalNumberOfRows());
				try {
			
			
			WebElement Search = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@type='search']")));
			Search.sendKeys(DEB_id);
			 
			
			WebElement universityname = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("(//tbody[@id='tbodyDisplay']//tr//td[7])[1]")));
			String universityName= universityname.getText();
			
			WebElement Coursename = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("(//tbody[@id='tbodyDisplay']//tr//td[8])[1]")));
			String CourseName=Coursename.getText();
			
		
			
			Thread.sleep(300);
			WebElement Search1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@type='search']")));
			Search1.clear();
			
			 
			   row1.createCell(0).setCellValue(DEB_id);
			   row1.createCell(1).setCellValue(universityName);
			   row1.createCell(2).setCellValue(CourseName);
			   
				}
				catch(Exception r)
				{
					String failed = DEB_id;	
					row1.createCell(0).setCellValue(failed);
					Thread.sleep(500);
					WebElement Search1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@type='search']")));
					Search1.clear();
				}
			fileOut = new FileOutputStream("D://Automation_data//SVU_Admission_20250401_4.xlsx");
	   	    workbook.write(fileOut);
	   	 Thread.sleep(500);  
	   	
		} 	 
}}
        	catch (Exception e) {
        		System.out.println( "No matching records found");
        	
		}
        finally {
            try {
            if (fileOut != null) {
                fileOut.close();
            };
            if (fis != null) {
                fis.close();
            }
            
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Close workbook after operations are done
        // Note: Don't close the workbook until all operations are finished
        if (driver != null) {
         driver.quit(); // Close the WebDriver session
          
        }
	}
	} 	

private String getCellValue(Cell cell) {
    if (cell == null) {
        return "";
    }
    switch (cell.getCellType()) {
        case STRING:
            return cell.getStringCellValue();
        case NUMERIC:
            if (DateUtil.isCellDateFormatted(cell)) {
                // If the cell contains a date, convert it to a string in the desired format
                java.util.Date date = cell.getDateCellValue();
                SimpleDateFormat sdf = new SimpleDateFormat("dd-MMM-yyyy"); // Customize format
                return sdf.format(date);
            } else {
                // Handle numeric values as needed
                return String.valueOf(cell.getNumericCellValue());
            }
        case BOOLEAN:
            return String.valueOf(cell.getBooleanCellValue());
        default:
            return "";
    }
}
private boolean isElementClickable(WebDriver driver, WebElement element) {
    try {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(1));
        wait.until(ExpectedConditions.elementToBeClickable(element));
        return true;
    } catch (Exception e) {
        return false;
    }
}
private void switchToNewWindow(WebDriver driver) {
Set<String> windowHandles = driver.getWindowHandles();
for (String windowHandle : windowHandles) {
    driver.switchTo().window(windowHandle);
}}
}
