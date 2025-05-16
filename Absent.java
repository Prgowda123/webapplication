package ksp;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class Absent {
	
@Test
public void sample() throws IOException, InterruptedException
{
	
//	driver.get("https://examcentre.svuddeonline.in/login");
//	WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
//	
	FileInputStream fis1 = new FileInputStream("D://Automation_data//absent.xlsx");//"D:\steno\TestDataAPC.xlsx"
	XSSFWorkbook workbook = new XSSFWorkbook(fis1);
	Sheet sheet = workbook.getSheetAt(0);

	// Select select = null;
	int rowCount = sheet.getPhysicalNumberOfRows();

	// Loop through rows in the Excel sheet
	// int rowCount = sheet.getPhysicalNumberOfRows();

	
	
	
	for (int i = 24; i <=32; i++) { // Start from row 1 to skip header
		Row row = sheet.getRow(i);
		
		if (row == null) {
			System.out.println("Skipping empty row: " + i);
			continue;
		}
        try {
		if (row != null) {
			ChromeDriver driver = new ChromeDriver();
	
			driver.manage().window().maximize();
			driver.get("https://examcentre.svuddeonline.in/login");
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
			String Applicant_id = getCellValue(row.getCell(1));
			String  Applicant_IdentityCard = getCellValue(row.getCell(2));
			
			WebElement Applicant_FullName = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Login_UserName")));
			Applicant_FullName.sendKeys(Applicant_id);
			double doubleValue1 = Double.parseDouble(Applicant_IdentityCard.trim()); // Handles "625.0"
			int intValue1 = (int) doubleValue1; // Converts to 625
			String intAsString1 = String.valueOf(intValue1);
			WebElement Father = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Login_Password")));
			Father.sendKeys(intAsString1);
			Thread.sleep(1000);
			
			WebElement login = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div//button")));
			login.click();
			Thread.sleep(1000);
			switchToNewWindow(driver);

			WebElement ApplyingType = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath ("//select[@id='selStatement']")));
			Select s=new Select(ApplyingType);
			s.selectByIndex(1);
			
			WebElement prioriti = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//select[@id='selExamDate']")));
			Select p=new Select(prioriti);
			p.selectByIndex(1);

			WebElement prioriti1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//select[@id='selSession']")));
			Select p1=new Select(prioriti1);
			p1.selectByIndex(1);
             Thread.sleep(1000);
			
			WebElement download = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div//button")));
			download.click();
			Thread.sleep(1000);
			
			 Alert a = driver.switchTo().alert();
			    a.accept();
			    Thread.sleep(1000);
			    
			   try { 
			    Alert alert = driver.switchTo().alert();
				alert.accept();
				System.out.println(i + " Sequence contains no elements");
			   }
			   catch (Exception e) {
				// TODO: handle exception
			}
			   driver.quit();
}
        }	catch (Exception e) {
			System.out.println(i + " Failed due to an error in username or password");
			 
		}
}}
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
