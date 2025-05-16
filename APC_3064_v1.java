package ksp;

import java.awt.AWTException;
import java.awt.Desktop;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;

import java.util.ArrayList;
import java.util.Date;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chromium.ChromiumDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class APC_3064_v1 {
	private ChromiumDriver driver;

	@Test
	public void runAutomation() throws IOException, InterruptedException, AWTException {

		try {
			WebDriver driver = new ChromeDriver();
			driver.manage().window().maximize();
			driver.get("https://ksp-recruitment.in/");
			Thread.sleep(3000);
			JavascriptExecutor jss = (JavascriptExecutor) driver;
			jss.executeScript("window.scrollBy(0,500)", "");
			Thread.sleep(3000);
			driver.findElement(By.linkText("CLICK HERE TO GO TO RECRUITMENT - 2022")).click();
			Thread.sleep(4000);

			Set<String> allwh = driver.getWindowHandles();
			for (String wh : allwh) {
				driver.switchTo().window(wh);
			}
			Thread.sleep(4000);
			JavascriptExecutor jss1 = (JavascriptExecutor) driver;
			jss.executeScript("window.scrollBy(0,1500)", "");
			Thread.sleep(3000);
			driver.findElement(By.xpath(
					"//td[text()=' Armed Police Constable (Male & Male Transgender) (CAR/DAR)-2022 ']/following-sibling::td[5]"))
					.click();
			Thread.sleep(4000);
			Set<String> allwh1 = driver.getWindowHandles();
			int count = allwh1.size();
			System.out.println(count);
			ArrayList<String> lan = new ArrayList<String>(allwh1);
			for (int i = 0; i < count; i++) {
				String k = lan.get(2);
				driver.switchTo().window(k);
			}
			Thread.sleep(3000);
			String parentWindow = driver.getWindowHandle();

			// Get all window handles
			Set<String> allWindows = driver.getWindowHandles();
			System.out.println(allWindows.size());
			allWindows.remove(parentWindow);
			System.out.println(allWindows);
			for (String allw : allWindows) {
				driver.switchTo().window(allw);
				driver.close();
			}
			driver.switchTo().window(parentWindow);

			driver.findElement(By.xpath("//a[@class='nav-menu nav-myapp-menu login-btn p-2']")).click();
			Thread.sleep(4000);

			// Handle new window - switch only once
			String mainWindowHandle = driver.getWindowHandle();
			Set<String> windowHandles = driver.getWindowHandles();
			for (String handle : windowHandles) {
				if (!handle.equals(mainWindowHandle)) {
					driver.switchTo().window(handle);
					break;
				}
			}

			// Load the Excel file
			FileInputStream fis = new FileInputStream("D://Automation_data//Apc_3064.xlsx");// C://Users//pallavi//eclipse-workspace//project//Book1.xlsx
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			Sheet sheet = workbook.getSheetAt(0);
			int rowcount = sheet.getPhysicalNumberOfRows();

			for (int i = 1; i < rowcount; i++) {
				Row row = sheet.getRow(i);
				if (row != null) {
					// Get ApplicationNo and DOB from each row
					String applicationNo = getCellValue(row.getCell(0)); // Applicant ID
					String dob = getCellValue(row.getCell(1)); // Date of Birth

					// Debugging or verification output
					System.out.println("Application No: " + applicationNo);
					System.out.println("DOB: " + dob);

					// Enter ApplicationNo and DOB in fields
					driver.findElement(By.id("Login_ApplicantId")).sendKeys(applicationNo);
					Thread.sleep(500); // Adjust sleep time as needed
					driver.findElement(By.id("Login_DateOfBirth")).sendKeys(dob);
					Thread.sleep(1000);

					// Submit the form by pressing Enter with Robot or by clicking a submit button
					Robot r = new Robot();
					r.keyPress(KeyEvent.VK_ENTER);
					r.keyRelease(KeyEvent.VK_ENTER);
					Thread.sleep(1000);

					JavascriptExecutor jss3 = (JavascriptExecutor) driver;
					jss.executeScript("window.scrollBy(0,500)", "");
					Thread.sleep(1000);
					WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
					Sheet sheet1 = workbook.getSheetAt(1);
					Row row1 = sheet1.createRow(sheet1.getPhysicalNumberOfRows());
					WebElement appno = wait.until(ExpectedConditions.presenceOfElementLocated(
							By.xpath("//td[contains(text(),' Application No. ')]//../td[2]")));
					String Applno = appno.getText();
					row1.createCell(0).setCellValue(Applno);

					WebElement DOB = wait.until(ExpectedConditions
							.presenceOfElementLocated(By.xpath("//td[contains(text(),' Date of Birth')]//../td[2]")));
					String Dateofbirth = DOB.getText();
					row1.createCell(1).setCellValue(Dateofbirth);

					WebElement written = wait.until(ExpectedConditions
							.presenceOfElementLocated(By.xpath("//h5[contains(text(),'Written Exam')]")));
					String writenexam = written.getText();
					System.out.println(writenexam);

					if (isElementClickable(driver, written)) {
						WebElement examdate = wait.until(ExpectedConditions.presenceOfElementLocated(
								By.xpath("//td[contains(text(),'WRITTEN EXAM DATE' )]/following-sibling::td[1]")));
						String written_exam = examdate.getText();
						System.out.println(written_exam);
						row1.createCell(2).setCellValue(written_exam);
					} else {
						String exam = "WRITTEN EXAM IS NOT COMPLETED";
						row1.createCell(2).setCellValue(exam);
					}

					try {
						WebElement ENDURANCE = wait.until(ExpectedConditions.presenceOfElementLocated(
								By.xpath("//h5[contains(text(),'ENDURANCE TEST & PHYSICAL STANDARD TEST' )]")));
						String ENDURANCEexam = ENDURANCE.getText();
						System.out.println(ENDURANCEexam);
						jss.executeScript("window.scrollBy(0,500)", "");

						if (isElementClickable(driver, ENDURANCE)) {
							WebElement ENDURANCETEST = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(
									"//h5[contains(text(),'ENDURANCE TEST & PHYSICAL STANDARD TEST' )]/../..//div/table/tbody/tr[1]/th/following-sibling::td[2]")));
							String ENDURANCE_TEST = ENDURANCETEST.getText();
							System.out.println(ENDURANCE_TEST);
							row1.createCell(3).setCellValue(ENDURANCE_TEST);
						}
					} catch (Exception e) {
						System.out.println("ETPST is COMPLETE");
					}
					try {
						WebElement MEDICAL = wait.until(ExpectedConditions
								.presenceOfElementLocated(By.xpath("//h5[contains(text(),' MEDICAL EXAM' )]")));
						String MEDICALexam = MEDICAL.getText();
						System.out.println(MEDICALexam);

						if (isElementClickable(driver, MEDICAL)) {
							WebElement DOC = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(
									"//td[text()=' DOCUMENT VERIFICATION DATE & TIME. ']/following-sibling::td[1]")));
							String Doc = DOC.getText();
							System.out.println(Doc);
							row1.createCell(4).setCellValue(Doc);
						} else {
							String DVexam = "DV not COMPLETED";
							row1.createCell(4).setCellValue(DVexam);
						}
					} catch (Exception e) {
						System.out.println("DV is not complete");
					}
					jss.executeScript("window.scrollBy(0,500)", "");
					try {
						WebElement MEDICAL = wait.until(ExpectedConditions
								.presenceOfElementLocated(By.xpath("//h5[contains(text(),' MEDICAL EXAM' )]")));

						String MEDICALexam = MEDICAL.getText();
						if (isElementClickable(driver, MEDICAL))

						{
							WebElement MED = wait.until(ExpectedConditions.presenceOfElementLocated(
									By.xpath("//td[text()=' MEDICAL EXAM DATE & TIME. ']/following-sibling::td[1]")));
							String Medical = MED.getText();
							System.out.println(Medical);
							row1.createCell(5).setCellValue(Medical);
						} else {
							String MVexam = "MV not COMPLETED";
							System.out.println(MVexam);
							row1.createCell(5).setCellValue(MVexam);
						}
					} catch (Exception e) {
						// TODO: handle exception
					}
					try {
						String st = "Office of the Superintendent Of Police, BIDAR District";

						WebElement ele = driver.findElement(
								By.xpath("//td[text()=' DOCUMENT VERIFICATION VENUE. ']/following-sibling::td[1]"));
						String title = ele.getText();
						System.out.println(title);
						if (st.equals(title)) {
							System.out.println("Both the venues are same");
						}
					} catch (Exception e) {
						// TODO: handle exception
					} // Wait for the submission to process
					driver.close();
					Thread.sleep(2000);
					driver.switchTo().window(mainWindowHandle);
					driver.findElement(By.xpath("//a[@class='nav-menu nav-myapp-menu login-btn p-2']")).click();
					String mainWindowHandle1 = driver.getWindowHandle();
					Set<String> windowHandles1 = driver.getWindowHandles();
					Thread.sleep(2000);

					for (String handle1 : windowHandles1) {
						if (!handle1.equals(mainWindowHandle1)) {
							driver.switchTo().window(handle1);
							break;
						}
					}
					FileOutputStream file = new FileOutputStream("D://Automation_data//Apc_3064.xlsx");
					workbook.write(file);
					file.close();
				}
			}

			
			 // ✅ Auto-open Excel file
            try {
                File excelFile = new File("D://Automation_data//Apc_3064.xlsx");
                if (excelFile.exists()) {
                    if (Desktop.isDesktopSupported()) {
                        Desktop.getDesktop().open(excelFile);
                    } else {
                        System.out.println("Desktop not supported.");
                    }
                } else {
                    System.out.println("Excel file not found.");
                }
            } catch (IOException e) {
                System.out.println("Could not open Excel file.");
                e.printStackTrace();
            }
        
    
			// Close resources
			workbook.close();
			fis.close();
			driver.quit();

		} catch (IOException e) {
			e.printStackTrace();

		} finally {
			// Close the browser at the end
			// driver.quit();
		}
	}

	// Helper method to handle different cell types
	private static String getCellValue(Cell cell) {
		if (cell == null) {
			return "";
		}

		switch (cell.getCellType()) {
		case NUMERIC:
			// Check if the numeric cell is a date
			if (DateUtil.isCellDateFormatted(cell)) {
				Date date = cell.getDateCellValue();
				SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
				return dateFormat.format(date);
			} else {
				// Return numeric value as a String (cast to long to avoid decimals)
				return String.valueOf((long) cell.getNumericCellValue());
			}
		case STRING:
			return cell.getStringCellValue();
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

}
