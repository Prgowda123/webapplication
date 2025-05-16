package apc_nkk;

import java.awt.AWTException;
import java.awt.Robot;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Set;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class MY_application {
@Test
public void sample() throws AWTException, InterruptedException, IOException
{
	ChromeDriver driver = new ChromeDriver();
	driver.manage().window().maximize();
	JavascriptExecutor jss = (JavascriptExecutor) driver;
	driver.get("http://172.10.1.159:9017");
	WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
	Actions act = new Actions(driver);
	Robot r = new Robot();
	wait.until(ExpectedConditions.presenceOfElementLocated(By.linkText("My Application"))).click();
	switchToNewWindow(driver);
	WebElement ApplicationLogin = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h2[contains(text(),'Application Login')]")));
	String Application_Login = ApplicationLogin.getText();
	
	WebElement ApplicationNumber = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//label[contains(text(),'Application Number')]")));
	String Application_Number = ApplicationNumber.getText();
	
	WebElement date = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//label[contains(text(),'Date of Birth')]")));
	String dob = date.getText();
	
	WebElement forgot = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a[contains(text(),'Forgot Application No?')]")));
	String forgotapp = forgot.getText();
	
	WebElement login = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[contains(text(),'Login')]")));
	String log = login.getText();
	
	
	WebElement appno = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ApplicantModel_ApplicationNo")));
	appno.sendKeys("0000046");

	WebElement dateofbirth = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ApplicantModel_DateOfBirth")));
	dateofbirth.sendKeys("22/11/1999");
	driver.findElement(By.id("login-submit")).click();
	Thread.sleep(3000);
	switchToNewWindow(driver);
	jss.executeScript("window.scrollBy(0,500)", "");
	
	WebElement MyApplication = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[contains(text(),'My Application')]")));
	String My_Application = MyApplication.getText();
	
	WebElement Applicant = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h5[contains(text(),'Applicant Photo')]")));
	String Applicantphoto = Applicant.getText();
	
	WebElement Signature = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h5[contains(text(),'Signature')]")));
	String signature = Signature.getText();
	
	WebElement Idcard = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h5[contains(text(),'ID Card')]")));
	String idcard = Idcard.getText();
	
	WebElement ThumbImpression = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h5[contains(text(),'Thumb Impression photo')]")));
	String Thumb_Impression = ThumbImpression.getText();
	
	WebElement ApplicantDetails = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h4[contains(text(),'Applicant Details')]")));
	String ApplicantDetail = ApplicantDetails.getText();
	
	
	WebElement applicant = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),'Applicant Full Name')]")));
	String applicantname = applicant.getText();
	
	WebElement applying = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),'Applying as')]")));
	String applyingas = applying.getText();
	
	WebElement dateof = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),'Date of Birth')]")));
	String dateofbirt = dateof.getText();
	
	WebElement gen = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),'Gender')]")));
	String Gen = gen.getText();
	
	WebElement adhar = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),'Aadhar Number')]")));
	String adharno = adhar.getText();
	
	WebElement Category = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),'Category')]")));
	String category = Category.getText();
	
	WebElement email = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),'Email Id')]")));
	String emailid = email.getText();
	
	WebElement Mobile = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),'Mobile Number')]")));
	String Mobileno = Mobile.getText();
	
	WebElement Unit = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//th[contains(text(),'Unit')]")));
	String unit = Unit.getText();
	
	WebElement DOWNLOAD = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[contains(text(),'DOWNLOAD APPLICATION')]")));
	String DOWNLOADapp = DOWNLOAD.getText();
	
	FileInputStream fis1 = new FileInputStream("D://APC//APC_NKK.xlsx");
	XSSFWorkbook workbook = new XSSFWorkbook(fis1);
    XSSFSheet sheet = workbook.getSheetAt(1);
	int rownumber = 9;
	XSSFRow row1 = sheet.createRow(rownumber);
	
	row1.createCell(1).setCellValue(Application_Login);
	 row1.createCell(2).setCellValue(Application_Number);
	 row1.createCell(3).setCellValue(dob);
	 row1.createCell(4).setCellValue(forgotapp);
	 row1.createCell(5).setCellValue(log);
	 row1.createCell(6).setCellValue(My_Application);
	 row1.createCell(7).setCellValue(Applicantphoto);
	 row1.createCell(8).setCellValue(signature);
	 row1.createCell(9).setCellValue(idcard);
	 row1.createCell(10).setCellValue(Thumb_Impression);
	 row1.createCell(11).setCellValue(ApplicantDetail);
	 row1.createCell(12).setCellValue(applicantname);
	 row1.createCell(13).setCellValue(applyingas);
	 row1.createCell(14).setCellValue(dateofbirt);
	 row1.createCell(15).setCellValue(Gen);
	 row1.createCell(16).setCellValue(adharno);
	 row1.createCell(17).setCellValue(category);
	 row1.createCell(18).setCellValue(emailid);
	 row1.createCell(19).setCellValue(Mobileno);
	 row1.createCell(20).setCellValue(unit);
	 row1.createCell(21).setCellValue(DOWNLOADapp);

	 
	 FileOutputStream file = new FileOutputStream("D://Automation_data//TestDataAPC_NKK.xlsx");
	 workbook.write(file);
	 file.close();
	 driver.quit();
	
	
}

private void switchToNewWindow(WebDriver driver)
{
	Set<String> window = driver.getWindowHandles();
	for (String all : window) {
		driver.switchTo().window(all);
	}
}
}
