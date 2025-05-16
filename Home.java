package apc_kk;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class Home {
@Test
public void sample() throws IOException
{
	ChromeDriver driver = new ChromeDriver();
	driver.manage().window().maximize();
	driver.get("http://172.10.1.159:9016");
	JavascriptExecutor jss = (JavascriptExecutor) driver;
	WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
	
	WebElement home = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//header//div//li[1]")));
	String Home = home.getText();
    
	WebElement Brouchure = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//header//div//li[2]")));
	String Brouchur = Brouchure.getText();

	WebElement Notification = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//header//div//li[3]")));
	String notification = Notification.getText();
	
	WebElement NewApplication = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//header//div//li[4]")));
	String New_Application = NewApplication.getText();
	
	WebElement MyApplication = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//header//div//li[5]")));
	String My_Application = MyApplication.getText();
	
	WebElement RECRUITMENT = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h5[contains(text(),'NOTIFICATION NO. - XY / RECRUITMENT - XY / 2023-24')]")));
	String RECRUITMENTx = RECRUITMENT.getText();
	
	WebElement rECRUITMENT = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h5[contains(text(),'ಅಧಿಸೂಚನೆ ಸಂಖ್ಯೆ - XY / ನೇಮಕಾತಿ - XY / 2023-24')]")));
	String rECRUITMENTs = rECRUITMENT.getText();
	
	WebElement CONSTABLE = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h3[contains(text(),'ARMED POLICE CONSTABLE (MALE & MALE TRANSGENDER) (KK) (CAR/DAR)-2023 (XYZ POSTS)')]")));
	String CONSTABLEs = CONSTABLE.getText();
	
	WebElement ARMED = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h4[contains(text(),'ಸಶಸ್ತ್ರ ಪೋಲೀಸ್ ಕಾನ್ಸ್ಟೇಬಲ್ ')]")));
	String ARME = ARMED.getText();
	
	WebElement Apply = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a[contains(text(),'Apply Now')]")));
	String Applynow = Apply.getText();
	
	WebElement Download = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a[contains(text(),'Download Notification')]")));
	String Downloadnotification = Download.getText();
	
	WebElement News = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h2[contains(text(),' News & Events')]")));
	String New = News.getText();
	jss.executeScript("arguments[0].scrollIntoView(true);", News);
	
	WebElement KeyDates = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h2[contains(text(),' Key Dates')]")));
	String Key_Dates = KeyDates.getText();
	
	WebElement sl = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),'SL.NO')]")));
	String slno = sl.getText();
	jss.executeScript("arguments[0].scrollIntoView(true);", sl);
	WebElement DESCRIPTION = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),'DESCRIPTION ')]")));
	String DESCRIPTIONs = DESCRIPTION.getText();
	
	WebElement date = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),'DATES')]")));
	String dates = date.getText();
	
	WebElement AgeCriteria = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h2[contains(text(),' Age Criteria')]")));
	String Age_Criteria = AgeCriteria.getText();
	
	WebElement CATEGORY = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),' CATEGORY ')]")));
	String CATEGORy = CATEGORY.getText();
	
	WebElement Minimum = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),'Minimum Age ')]")));
	String Minimumage = Minimum.getText();
	
	WebElement age = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),' AGE')]")));
	String Age = age.getText();
	
	WebElement DATE = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),' DATE ')]")));
	String DATEs = DATE.getText();
	
	WebElement Maximum = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),'Maximum Age')]")));
	String Maximumage = Maximum.getText();
	
	WebElement OVERAGE = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),' OVERAGE ')]")));
	String OVERAGEe = OVERAGE.getText();
	
	WebElement UNDERAGE = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),' UNDERAGE ')]")));
	String UNDERAGEe = UNDERAGE.getText();
	
	WebElement Know = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[contains(text(),'Know your')]")));
	String Knowage = Know.getText();
	
	WebElement Fees = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h2[contains(text(),' Fees Details')]")));
	String Fee = Fees.getText();
	jss.executeScript("arguments[0].scrollIntoView(true);", Fees);

	
	WebElement NOTE = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h4[contains(text(),'NOTES:')]")));
	String NOTEs = NOTE.getText();
	
	FileInputStream fis1 = new FileInputStream("D://Automation_data//TestDataAPC_KK.xlsx");
	XSSFWorkbook workbook = new XSSFWorkbook(fis1);
	XSSFSheet sheet = workbook.getSheetAt(2);
	int rownumber = 15;
	XSSFRow row1 = sheet.createRow(rownumber);
	
	row1.createCell(1).setCellValue(Home);
	 row1.createCell(2).setCellValue(Brouchur);
	 row1.createCell(3).setCellValue(notification);
	 row1.createCell(4).setCellValue(New_Application);
	 row1.createCell(5).setCellValue(My_Application);
	 row1.createCell(6).setCellValue(RECRUITMENTx);
	 row1.createCell(7).setCellValue(rECRUITMENTs);
	 row1.createCell(8).setCellValue(CONSTABLEs);
	 row1.createCell(9).setCellValue(ARME);
	 row1.createCell(10).setCellValue(Applynow);
	 row1.createCell(11).setCellValue(Downloadnotification);
	 row1.createCell(12).setCellValue(New);
	 row1.createCell(13).setCellValue(Key_Dates);
	 row1.createCell(14).setCellValue(slno);
	 row1.createCell(15).setCellValue(DESCRIPTIONs);
	 row1.createCell(16).setCellValue(dates);
	 row1.createCell(17).setCellValue(Age_Criteria);
	 row1.createCell(18).setCellValue(CATEGORy);
	 row1.createCell(19).setCellValue(Minimumage);
	 row1.createCell(20).setCellValue(Age);
	 row1.createCell(21).setCellValue(DATEs);
	 row1.createCell(22).setCellValue(Maximumage);
	 row1.createCell(23).setCellValue(Age);
	 row1.createCell(24).setCellValue(DATEs);
	 row1.createCell(25).setCellValue(OVERAGEe);
	 row1.createCell(26).setCellValue(UNDERAGEe);
	 row1.createCell(27).setCellValue(Knowage);
	 row1.createCell(28).setCellValue(Fee);
	 row1.createCell(29).setCellValue(NOTEs);
	
	 FileOutputStream file = new FileOutputStream("D://Automation_data//TestDataAPC_KK.xlsx");
	 workbook.write(file);
	 file.close();
	 driver.quit();
	
}

}
