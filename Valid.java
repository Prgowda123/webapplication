

package Steno;

import org.testng.annotations.Test;
import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Set;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.testng.Reporter;
import org.testng.annotations.Test;

public class Valid
{

	
    @Test
    public void sample() throws InterruptedException, IOException, AWTException {
        WebDriver driver = new ChromeDriver();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        driver.get("http://172.10.1.159:9013");
        driver.manage().window().maximize();
        
        // Navigate to "New Application"
        driver.findElement(By.linkText("New Application")).click();
        String Mainwindow = driver.getWindowHandle();

        // Switch to the new window
        Set<String> allWindows = driver.getWindowHandles();
        for (String window : allWindows) {
            driver.switchTo().window(window);
        }

        // Scroll and interact with elements
        JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript("window.scrollBy(0,1500);");
        Thread.sleep(1000);

        driver.findElement(By.xpath("//input[1]")).click();
        Thread.sleep(500);

        driver.findElement(By.id("nextBtn")).click();

        allWindows = driver.getWindowHandles();
        for (String window : allWindows) {
            driver.switchTo().window(window);
        }
        Thread.sleep(1000);

        js.executeScript("window.scrollBy(0,400);");
      

        // Open Excel file for reading and writing
        FileInputStream fis = null;
        FileOutputStream fileOut = null;

        try {
			// Reading Excel File
			FileInputStream fis1 = new FileInputStream("D://Automation_data//TestData (2).xlsx");//"D:\steno\TestData (2).xlsx"
			XSSFWorkbook workbook = new XSSFWorkbook(fis1);
			Sheet sheet = workbook.getSheetAt(0);

			// Select select = null;
			int rowCount = sheet.getPhysicalNumberOfRows();

			// Loop through rows in the Excel sheet
			// int rowCount = sheet.getPhysicalNumberOfRows();

			
			Actions act = new Actions(driver);
			
			for (int i = 8; i <= 10; i++) { // Start from row 1 to skip header
				Row row = sheet.getRow(i);
				
				if (row == null) {
					System.out.println("Skipping empty row: " + i);
					continue;
				}

				if (row != null) {
				

					String name = getCellValue(row.getCell(2));
					String father = getCellValue(row.getCell(3));
					String mother = getCellValue(row.getCell(4));
					String email = getCellValue(row.getCell(5));
					String mob = getCellValue(row.getCell(6));
					String adhar = getCellValue(row.getCell(7));
					String dob = getCellValue(row.getCell(8));
					String doorno = getCellValue(row.getCell(11));
					String street = getCellValue(row.getCell(12));
					
					String taluk = getCellValue(row.getCell(13));
					String city = getCellValue(row.getCell(14));
					int dropdownIndex2 = (int) row.getCell(16).getNumericCellValue();
					String othdis = getCellValue(row.getCell(17));
					String pincode = getCellValue(row.getCell(18));
					String landmark = getCellValue(row.getCell(19));
					String Ndistrict = getCellValue(row.getCell(20));
					int caste = (int) row.getCell(31).getNumericCellValue();
					String subcaste = getCellValue(row.getCell(32));
					String issue = getCellValue(row.getCell(33));
					String govtjoin = getCellValue(row.getCell(35));
					String dept = getCellValue(row.getCell(36));
					String designation = getCellValue(row.getCell(41));
					int year_index = (int) row.getCell(38).getNumericCellValue();
					int month_index = (int) row.getCell(39).getNumericCellValue();
					int day_index = (int) row.getCell(40).getNumericCellValue();
					String dep_eq = getCellValue(row.getCell(43));
					String crime_detail = getCellValue(row.getCell(45));
					String CA_detail = getCellValue(row.getCell(47));
					int board1 = (int) row.getCell(49).getNumericCellValue();
					String othersslc = getCellValue(row.getCell(50));

					int paper1 = (int) row.getCell(51).getNumericCellValue();
					int sslcyear = (int) row.getCell(52).getNumericCellValue();
					String m_mark = getCellValue(row.getCell(54));
					String ob_mark = getCellValue(row.getCell(55));
					String reg_no = getCellValue(row.getCell(59));
					int puboard1 = (int) row.getCell(61).getNumericCellValue();
					int pucyear = (int) row.getCell(63).getNumericCellValue();
					String pugrade = getCellValue(row.getCell(68));
					String cgpa = getCellValue(row.getCell(69));
					String pucreg = getCellValue(row.getCell(70));
					int typist1 = (int) row.getCell(73).getNumericCellValue();
					int kannada = (int) row.getCell(75).getNumericCellValue();
					int adharid = (int) row.getCell(76).getNumericCellValue();
					String idco = getCellValue(row.getCell(77));
					String idmk1 = getCellValue(row.getCell(78));
					String idmk2 = getCellValue(row.getCell(79));
					String photo = getCellValue(row.getCell(80));
					String thumb1 = getCellValue(row.getCell(81));
					String idcard = getCellValue(row.getCell(82));
					int applyingpost = (int) row.getCell(1).getNumericCellValue();
					String doornop = getCellValue(row.getCell(22));
					String streetp = getCellValue(row.getCell(23));
					String landmarkp = getCellValue(row.getCell(30));
					String talukp = getCellValue(row.getCell(24));
					String cityp = getCellValue(row.getCell(25));
					int dropdownIndex5 = (int) row.getCell(27).getNumericCellValue();
					String pincodep = getCellValue(row.getCell(29));
					int dropdownIndex6 = (int) row.getCell(26).getNumericCellValue();
					String othboard = getCellValue(row.getCell(62));



				//	System.out.println("max mark Number from Excel: " + m_mark);
				//	System.out.println("obtain from Excel: " + ob_mark);

				//	System.out.println("Mobile Number from Excel: " + mob);
				//	System.out.println("DOB from Excel: " + dob);
				

				//	System.out.println("Application No: " + name);
				//	System.out.println("Father's Name: " + father);
					
				   // System.out.println(applyingpost);
				    

					// Fill the form
			        driver.manage().window().maximize();

					WebElement dropdown = driver.findElement(By.id("Applicant_PostUnitCode"));
					Select select1 = new Select(dropdown);
					select1.selectByIndex(applyingpost);

					wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_FullName"))).sendKeys(name);

					wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_FatherName"))).sendKeys(father);
				
					wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_MotherName"))).sendKeys(mother);
					
					wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EmailId"))).sendKeys(email);
				
					wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_MobileNo"))).sendKeys(mob);
			
					wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_AadharNo"))).sendKeys(adhar);

					// Enter Date of Birth
					js.executeScript("window.scrollBy(0,500);");
					WebElement dobField = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_DateOfBirth")));
					dobField.sendKeys(dob);
					Thread.sleep(200);
					Robot robot = new Robot();
					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER);
					js.executeScript("window.scrollBy(0,200)", "");

					WebElement gen = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GenderCode")));
					Select ss1 = new Select(gen);
					Thread.sleep(100);
				
					int dropdownIndex = (int) row.getCell(10).getNumericCellValue();
					ss1.selectByIndex(dropdownIndex);

			       // Applicant_ContactAddress
					wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_DoorNo"))).sendKeys(doorno);
				
					wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_Street"))).sendKeys(street);
					
					js.executeScript("window.scrollBy(0,200)", "");
					wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_Taluk"))).sendKeys(taluk);
					
					wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_City"))).sendKeys(city);
					WebElement stt =wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_UnionStateCode")));
					act.moveToElement(stt).click().perform();
					Thread.sleep(500);
					Select ss2 = new Select(stt);
					
					String state = stt.getText();

					// Get the dropdown index value from the Excel row (assuming it's a numeric value)
					int dropdownIndex1 = (int) row.getCell(15).getNumericCellValue();

					// Add the condition to skip index 1
					if (dropdownIndex1 != 23) { // Only proceed if the index is not 1
					    ss2.selectByIndex(dropdownIndex1);  // Select the dropdown item by the given index
					 //   System.out.println("Selected index: " + dropdownIndex1);
					} else {
					    System.out.println("Skipping index 1.");
					}

					// Scroll the window (optional)
					js.executeScript("window.scrollBy(0,200)", "");
					
					// Applicant_Reservation_CategoryCode

					WebElement dis = driver.findElement(By.id("Applicant_ContactAddress_DistrictCode"));
					Thread.sleep(500);
					// dis.click();
					Thread.sleep(500);
					Select ss3 = new Select(dis);
					Thread.sleep(500);
					ss3.selectByIndex(dropdownIndex2);
					String district = dis.getText();
					if (district.equals("Other")) { // Only proceed if the index is not 1
					    ss3.selectByIndex(dropdownIndex2);  // Select the dropdown item by the given index
					//    System.out.println("Selected index: " + dropdownIndex2);
					    

						WebElement otherddis = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_OtherDistrictName")));

						wait.until(ExpectedConditions.visibilityOf(otherddis));
						wait.until(ExpectedConditions.elementToBeClickable(otherddis));
						otherddis.sendKeys(othdis);
					} else {
						
					}
					Thread.sleep(500);
					driver.findElement(By.id("pincodeInput")).sendKeys(pincode);
					Thread.sleep(500);
					driver.findElement(By.id("Applicant_ContactAddress_Landmark")).sendKeys(landmark);
					Thread.sleep(500);
					driver.findElement(By.id("Applicant_NativeDistrict")).sendKeys(Ndistrict);
					Thread.sleep(500);
					

					WebElement add = driver.findElement(By.xpath("//input[@id='Applicant_ContactAddress_IsPermanentAddressSame' and @value='False']"));
					act.moveToElement(add).click().perform();
					String addr = add.getText();
					Thread.sleep(500);
					driver.findElement(By.id("Applicant_PermanentAddress_DoorNo")).sendKeys(doornop);
					Thread.sleep(500);
					driver.findElement(By.id("Applicant_PermanentAddress_Street")).sendKeys(streetp);
					Thread.sleep(500);

					WebElement t = driver.findElement(By.id("Applicant_PermanentAddress_Taluk"));
					t.sendKeys(talukp);
					Thread.sleep(500);
					driver.findElement(By.id("Applicant_PermanentAddress_City")).sendKeys(cityp);
					//js.executeScript("arguments[0].scrollIntoView(true);", cityp);

					Thread.sleep(500);// Applicant_ContactAddress_UnionStateCode
					WebElement stat = driver.findElement(By.id("Applicant_PermanentAddress_UnionStateCode"));
			     //	js.executeScript("arguments[0].scrollIntoView(true);", stat);

					Thread.sleep(500);
					act.moveToElement(stat).click().perform();
					Thread.sleep(500);
					Select ss30 = new Select(stat);
					Thread.sleep(500);
					ss30.selectByIndex(dropdownIndex6);
					
					

						WebElement dist = driver.findElement(By.id("Applicant_PermanentAddress_DistrictCode"));
					Thread.sleep(500);
					act.moveToElement(dist).click().perform();
					Thread.sleep(500);
					Select ss31 = new Select(dist);
			
					Thread.sleep(500);
					ss31.selectByIndex(dropdownIndex5);
					
			

				   driver.findElement(By.id("Applicant_PermanentAddress_Pincode")).sendKeys(pincodep);
					Thread.sleep(500);

					driver.findElement(By.id("Applicant_PermanentAddress_Landmark")).sendKeys(landmarkp);
					Thread.sleep(500);
						js.executeScript("window.scrollBy(0,100)", "");// Applicant_Reservation_CategoryCode
					WebElement cast = driver.findElement(By.id("Applicant_Reservation_CategoryCode"));
					act.moveToElement(cast).click().perform();
					Thread.sleep(500);
					Select ss4 = new Select(cast);
					ss4.selectByIndex(caste);
					Robot r = new Robot();
					driver.findElement(By.id("Applicant_Reservation_SubCaste")).sendKeys(subcaste);
					Thread.sleep(500);
					WebElement cd = driver.findElement(By.id("Applicant_Reservation_CategoryCertificateIssuedDate"));
					cd.sendKeys(issue);
					r.keyPress(KeyEvent.VK_ENTER);
					r.keyRelease(KeyEvent.VK_ENTER);
					Thread.sleep(500);
					js.executeScript("arguments[0].scrollIntoView(true)", cd);

					js.executeScript("window.scrollBy(0,200)", "");
					
		            WebDriverWait wait55 = new WebDriverWait(driver, Duration.ofSeconds(20));
		            Thread.sleep(1000);
					WebElement cdd = wait55.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_AreYouAGovermentEmployee' and @value='True']")));
					wait55.until(ExpectedConditions.visibilityOf(cdd));
					wait55.until(ExpectedConditions.elementToBeClickable(cdd));
					cdd.click();
					js.executeScript("arguments[0].scrollIntoView(true);", cdd);
				
				//	WebElement cdd = driver.findElement(By.xpath("//input[@id='Applicant_Reservation_AreYouAGovermentEmployee' and @value='True']"));
					//act.moveToElement(cdd).click().perform();
					Thread.sleep(500);
					WebElement gv = driver.findElement(By.xpath("//input[@id='Applicant_Reservation_GovermentServiceDetail_JoiningDate']"));
					//js.executeScript("arguments[0].scrollIntoView(true);", gv);

					act.moveToElement(gv).click().perform();
					Thread.sleep(500);
					gv.sendKeys(govtjoin);
					r.keyPress(KeyEvent.VK_ENTER);
					r.keyRelease(KeyEvent.VK_ENTER);
					Thread.sleep(500);
					driver.findElement(By.id("Applicant_Reservation_GovermentServiceDetail_Department")).sendKeys(dept);
					Thread.sleep(500);//
					driver.findElement(By.id("Applicant_Reservation_GovermentServiceDetail_Designation")).sendKeys(designation);
					Thread.sleep(500);
					js.executeScript("window.scrollBy(0,500)", "");// Applicant_Reservation_CategoryCode
					WebElement year = driver.findElement(By.id("Applicant_Reservation_GovermentServiceDetail_YearsInService"));
					year.click();
					Thread.sleep(500);
					Select ss5 = new Select(year);
					ss5.selectByIndex(year_index);
					WebElement month = driver.findElement(By.id("Applicant_Reservation_GovermentServiceDetail_MonthsInService"));
					month.click();
					Thread.sleep(200);
					Select ss6 = new Select(month);
					ss6.selectByIndex(month_index);
					WebElement day = driver.findElement(By.id("Applicant_Reservation_GovermentServiceDetail_DaysInService"));
					day.click();
					Thread.sleep(500);
					Select ss7 = new Select(day);
					ss7.selectByIndex(day_index);
					Thread.sleep(500);
					WebElement ab =wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_GovermentServiceDetail_HasDepartmentEnquiry' and @value='True']")));
					//WebElement ab = driver.findElement(By.xpath("//input[@id='Applicant_CriminalActivity_HasDepartmentEnquiry' and @value='True']"));
					act.moveToElement(ab).click().perform();
					Thread.sleep(500);

					// Enter details in the "Department Enquiry Detail" field
					wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GovermentServiceDetail_DepartmentEnquiryDetail"))).sendKeys(dep_eq);

					js.executeScript("window.scrollBy(0,400)");
					Thread.sleep(500);

					WebElement cr = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_CriminalActivity_IsInvolvedInCriminalActivity' and @value='True']")));
					act.moveToElement(cr).click().perform();
					Thread.sleep(500);

					wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_CriminalActivity_CaseDetail"))).sendKeys(crime_detail);
					js.executeScript("window.scrollBy(0,200)", "");
				
					Thread.sleep(5000);
					WebElement crd = driver.findElement(By.xpath("//input[@id='Applicant_CriminalActivity_IsConvictedInCriminalCase' and @value='True']"));
					act.moveToElement(crd).click().perform(); // Move to the element and click it
					Thread.sleep(500);

					wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_CriminalActivity_ConvictionDetail"))).sendKeys(CA_detail);
					Thread.sleep(500);

					// for sslc yes
				
					WebElement sslc = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_IsSSLCHolder' and @value='True']")));
					Thread.sleep(300);
					act.moveToElement(sslc).click().perform();
					
					
					//for sslc board
					WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(20));
					WebElement board = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_QualificationBoardCode")));
					wait1.until(ExpectedConditions.visibilityOf(board));
					wait1.until(ExpectedConditions.elementToBeClickable(board));

					// Scroll to the element
					JavascriptExecutor js1 = (JavascriptExecutor) driver;
					Select ss8 = new Select(board);
					ss8.selectByIndex(board1);
			
					if (board1==3) {
					  
					    ss8.selectByIndex(board1); 
					
					  
					    WebElement sslcother = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_OtherBoardName")));
					    wait1.until(ExpectedConditions.visibilityOf(sslcother));
					    wait1.until(ExpectedConditions.elementToBeClickable(sslcother)); 
				        sslcother.sendKeys(othersslc);  

					} else {
					  //  System.out.println("Condition not met: board1 is not equal to 3.");
					}
					js.executeScript("window.scrollBy(0,200)");
					
				
					WebDriverWait wait2 = new WebDriverWait(driver, Duration.ofSeconds(20));
					WebElement paper = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_KannadaLanguagePaper")));
					wait2.until(ExpectedConditions.visibilityOf(paper));
					wait2.until(ExpectedConditions.elementToBeClickable(paper));

					// Scroll to the element
					if(applyingpost!=2) {
					JavascriptExecutor jss11 = (JavascriptExecutor) driver;
					//jss11.executeScript("arguments[0].scrollIntoView(true);", paper);
					}
					Select ss9 = new Select(paper);
					ss9.selectByIndex(paper1);

					
					// for sslc passing year
					WebDriverWait wait3 = new WebDriverWait(driver, Duration.ofSeconds(20));
					WebElement year3 = wait3.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_YearOfPassing")));
					wait3.until(ExpectedConditions.visibilityOf(year3));
					wait3.until(ExpectedConditions.elementToBeClickable(year3));

					// Scroll to the element
					JavascriptExecutor jss12 = (JavascriptExecutor) driver;
				//	jss12.executeScript("arguments[0].scrollIntoView(true);", year3);

					Select ss11 = new Select(year3);
					ss11.selectByIndex(sslcyear);
					
                    // for sslc mark or garde
					WebDriverWait wait4 = new WebDriverWait(driver, Duration.ofSeconds(20));
					WebElement mark = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_MarkType")));
					Thread.sleep(1000);
					act.moveToElement(mark).click().perform();

					// Parse m_mark as a double and convert to integer
					double doubleValue = Double.parseDouble(m_mark.trim()); // Handles "625.0"
					int intValue = (int) doubleValue; // Converts to 625

					// Convert the integer back to String for sendKeys
					String intAsString = String.valueOf(intValue);

					// Locate the input field
					WebElement maxmark = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_Score_Maximum")));

					
					wait1.until(ExpectedConditions.visibilityOf(maxmark));
					wait1.until(ExpectedConditions.elementToBeClickable(maxmark));

					// Clear existing data and send the converted integer
					maxmark.clear();
					maxmark.sendKeys(intAsString);
				//	System.out.println("Successfully entered: " + intAsString);

					// Parse m_mark as a double and convert to integer
					double doubleValue1 = Double.parseDouble(ob_mark.trim()); // Handles "625.0"
					int intValue1 = (int) doubleValue1; // Converts to 625

					// Convert the integer back to String for sendKeys
					String intAsString1 = String.valueOf(intValue1);

					WebElement omark = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_Score_Obtained")));
					wait1.until(ExpectedConditions.visibilityOf(omark));
					wait1.until(ExpectedConditions.elementToBeClickable(omark));
					omark.clear();
					omark.sendKeys(intAsString1);

				//	System.out.println("Successfully entered: " + intAsString1);

					WebDriverWait wait5 = new WebDriverWait(driver, Duration.ofSeconds(20));
					WebDriverWait wait51 = new WebDriverWait(driver, Duration.ofSeconds(20));

					WebElement regno = wait51.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@name='Applicant.EducationalQualification.SSLCQualification.RegistrationNo' and @id='Applicant_EducationalQualification_SSLCQualification_RegistrationNo']")));
					wait51.until(ExpectedConditions.visibilityOf(regno));
					wait51.until(ExpectedConditions.elementToBeClickable(regno));

					JavascriptExecutor js11 = (JavascriptExecutor) driver;
					js11.executeScript("arguments[0].scrollIntoView(true);", regno);
					// js11.executeScript("arguments[0].click();", regno); // Ensure it's focused

					regno.clear();
					regno.sendKeys(reg_no);
					js.executeScript("window.scrollBy(0,100)");
					
					
					// for puc
					WebElement puc = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_IsPUCHolder' and @value='True']")));
					Thread.sleep(500);
					act.moveToElement(puc).click().perform();
					
					//for puc board
					WebDriverWait wait11 = new WebDriverWait(driver, Duration.ofSeconds(20));
					WebElement puboard = wait11.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_QualificationBoardCode")));
					wait11.until(ExpectedConditions.visibilityOf(puboard));
					wait11.until(ExpectedConditions.elementToBeClickable(puboard));

					// Scroll to the element
				//	js.executeScript("arguments[0].scrollIntoView(true);", puboard);

					Select ss14 = new Select(puboard);
					ss14.selectByIndex(puboard1);
				//	System.out.println("Dropdown option selected!");

					if (puboard1==3) {
					   
					    ss14.selectByIndex(puboard1);  // Select the dropdown item by the given index
					  //  System.out.println("Selected index: " + puboard1);
					 
					    WebElement puother = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_OtherBoardName")));
					    wait1.until(ExpectedConditions.visibilityOf(puother));
					    wait1.until(ExpectedConditions.elementToBeClickable(puother));
					    Thread.sleep(1000); 
					    puother.sendKeys(othboard);
					   
					   
					    
					} else {
					   
					    // Log to make sure the field is ready
					 //   System.out.println("Field is ready: " );
					    }
					
					//for puc year
					WebDriverWait wait31 = new WebDriverWait(driver, Duration.ofSeconds(20));
					WebElement puyear3 = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_YearOfPassing")));
					wait31.until(ExpectedConditions.visibilityOf(puyear3));
					wait31.until(ExpectedConditions.elementToBeClickable(puyear3));

					// Scroll to the element
				//	jss12.executeScript("arguments[0].scrollIntoView(true);", puyear3);

					Select ss111 = new Select(puyear3);
					ss111.selectByIndex(pucyear);
					Thread.sleep(500);
					
					
                    // for puc garde or mark
					WebElement grade = wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_PUCQualification_MarkType' and @value='G']")));
					
					act.moveToElement(grade).click().perform();

					WebElement gr = wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_PUCQualification_Grade_Grade' and @name='Applicant.EducationalQualification.PUCQualification.Grade.Grade']")));
					wait1.until(ExpectedConditions.visibilityOf(gr));
					wait1.until(ExpectedConditions.elementToBeClickable(gr));
					gr.clear();
					gr.sendKeys(pugrade);
					WebElement per = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_ScorePercentage")));
					wait1.until(ExpectedConditions.visibilityOf(per));
					wait1.until(ExpectedConditions.elementToBeClickable(per));
					per.clear();
					per.sendKeys(cgpa);

					WebElement pureg = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_RegistrationNo")));
					wait1.until(ExpectedConditions.visibilityOf(pureg));
					wait1.until(ExpectedConditions.elementToBeClickable(pureg));
					pureg.clear();
					pureg.sendKeys(pucreg);

					// for degree
					WebElement deg = driver.findElement(By.xpath("//input[@id='Applicant_EducationalQualification_IsDegreeHolder' and @value='True']"));
					Thread.sleep(500);
					act.moveToElement(deg).click().perform();
					Thread.sleep(500);
					// jss12.executeScript("arguments[0].scrollIntoView(true);", deg);

					
					// for qualification
					WebElement senior = driver.findElement(By.xpath("//input[@id='Applicant_TypistAssistant_IsPassedInQualifyExam' and @value='True']"));
					Thread.sleep(500);
					act.moveToElement(senior).click().perform();
					

					WebDriverWait wait6 = new WebDriverWait(driver, Duration.ofSeconds(20));
					WebElement opt11 = wait6.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_TypistAssistant_QualificationCode")));
					wait6.until(ExpectedConditions.visibilityOf(opt11));
					wait6.until(ExpectedConditions.elementToBeClickable(opt11));
					Thread.sleep(500);
					Select ss15 = new Select(opt11);
					ss15.selectByIndex(typist1);
					String selectedOption = ss15.getFirstSelectedOption().getText();

					
					if (applyingpost != 2) {
						WebElement typist = driver.findElement(By.xpath("//input[@id='Applicant_StenographerAssistant_IsPassedInQualifyExam' and @value='True']"));
						Thread.sleep(500);
						act.moveToElement(typist).click().perform();
						

						WebDriverWait wait61 = new WebDriverWait(driver, Duration.ofSeconds(20));
						WebElement topt1 = wait61.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_StenographerAssistant_QualificationCode")));
						wait61.until(ExpectedConditions.visibilityOf(topt1));
						wait61.until(ExpectedConditions.elementToBeClickable(topt1));

						// Scroll to the dropdown element
						jss12.executeScript("arguments[0].scrollIntoView(true);", topt1);
						Thread.sleep(1000);
						Select ss16 = new Select(topt1);
						ss16.selectByIndex(kannada);
					}
					
					
					// for document upload
					
					WebDriverWait wait17 = new WebDriverWait(driver, Duration.ofSeconds(20));
					WebElement dopt1 = wait17.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_IdentityCardTypeCode")));
					wait17.until(ExpectedConditions.visibilityOf(dopt1));
					wait17.until(ExpectedConditions.elementToBeClickable(dopt1));

				
					Select ss17 = new Select(dopt1);
					ss17.selectByIndex(adharid);
					Thread.sleep(500);  
                 
					if (adharid != 1) {
					    // Only proceed if the index is not 1
					    ss17.selectByIndex(adharid);  // Select the dropdown item by the given index
					 //   System.out.println("Selected index: " + adharid);
					  //  / Wait for the text field to be present and interactable
					    WebElement idcod = wait17.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_UploadedIDNo")));
					    wait17.until(ExpectedConditions.visibilityOf(idcod));
					    wait17.until(ExpectedConditions.elementToBeClickable(idcod));
					  
					    idcod.sendKeys(idco);
					//    System.out.println("Field is ready: " + idcod.isEnabled());
					    
					} else {
					   
					    // Log to make sure the field is ready
					 //   System.out.println("Field is ready: " );
					   
					}   

					WebDriverWait wait20 = new WebDriverWait(driver, Duration.ofSeconds(20));
					WebDriverWait wait18 = new WebDriverWait(driver, Duration.ofSeconds(20));

					WebElement idtm = wait18.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_IdentificationMark_01")));
					wait18.until(ExpectedConditions.visibilityOf(idtm));
					wait18.until(ExpectedConditions.elementToBeClickable(idtm));
					idtm.clear();
					idtm.sendKeys(idmk1);

					WebDriverWait wait19 = new WebDriverWait(driver, Duration.ofSeconds(20));
					WebElement idtm2 = wait19.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_IdentificationMark_02")));
					wait19.until(ExpectedConditions.visibilityOf(idtm2));
					wait19.until(ExpectedConditions.elementToBeClickable(idtm2));
					idtm2.clear();
					idtm2.sendKeys(idmk2);
					if (applyingpost != 1) {
						js.executeScript("window.scrollBy(0,400)", "");
						Thread.sleep(500);
					}
					
					
					WebElement file = driver.findElement(By.name("Applicant.Photo"));
					file.sendKeys(photo);

					Thread.sleep(100);
					WebElement thumb = driver.findElement(By.name("Applicant.Thumb"));
					thumb.sendKeys(thumb1);

					Thread.sleep(500);
					WebElement id = driver.findElement(By.name("Applicant.IdentityCard"));
					id.sendKeys(idcard);
					Thread.sleep(500);
					js.executeScript("window.scrollBy(0,200)", "");
					Thread.sleep(1000);
					
				try {	
					// for click on preview
					WebElement pr = driver.findElement(By.id("preview-btn"));
					act.moveToElement(pr).click().perform();
					r.keyPress(KeyEvent.VK_ENTER);
					r.keyRelease(KeyEvent.VK_ENTER);
					Thread.sleep(1000);
					
					//for preview contents
					
					WebDriverWait wait23 = new WebDriverWait(driver, Duration.ofSeconds(8));
					WebElement appname = wait23.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='candidateTypePreview']")));
					wait23.until(ExpectedConditions.visibilityOf(appname));
					wait23.until(ExpectedConditions.elementToBeClickable(appname));
					String appnamep = appname.getText();
				//	System.out.println(appname.getText());
					jss12.executeScript("arguments[0].scrollIntoView(true);", appname);
					Sheet sheet2 = workbook.getSheetAt(3);
				    Row row1 = sheet2.createRow(sheet2.getPhysicalNumberOfRows()); 
					WebDriverWait wait24 = new WebDriverWait(driver, Duration.ofSeconds(20));

				
					WebElement canname = wait24.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='candidatenameTypePreview']")));
					wait24.until(ExpectedConditions.visibilityOf(canname));
					wait24.until(ExpectedConditions.elementToBeClickable(canname));
					String can = canname.getText();
			//		System.out.println(canname.getText());
				//	jss12.executeScript("arguments[0].scrollIntoView(true);", canname);
					Thread.sleep(500);
					
					WebDriverWait wait26 = new WebDriverWait(driver, Duration.ofSeconds(20));
					WebElement fathername = wait26.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='fatherNamePreview']")));
					wait26.until(ExpectedConditions.visibilityOf(fathername));
					wait26.until(ExpectedConditions.elementToBeClickable(fathername));
					String fathernamep = fathername.getText();
					//System.out.println(fathername.getText());
					
					WebDriverWait wait27 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement mothername = wait27.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='MotherNamePreview']")));
					wait27.until(ExpectedConditions.visibilityOf(mothername));
					wait27.until(ExpectedConditions.elementToBeClickable(mothername));
					String mothernamep = mothername.getText();
				//	System.out.println(mothername.getText());
						
					WebDriverWait wait28 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement emailid = wait28.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='emailPreview']")));
					wait28.until(ExpectedConditions.visibilityOf(emailid));
					wait28.until(ExpectedConditions.elementToBeClickable(emailid));
					String emailidp = emailid.getText();
				//	System.out.println(emailid.getText());
					
					WebDriverWait wait29 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement mobileno = wait29.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='MobileNoPreview']")));
					wait29.until(ExpectedConditions.visibilityOf(mobileno));
					wait29.until(ExpectedConditions.elementToBeClickable(mobileno));
					String mobilenop = mobileno.getText();
				//	System.out.println(mobileno.getText());
					
					WebDriverWait wait30 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement adharno = wait30.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='aadharPreview']")));
					wait30.until(ExpectedConditions.visibilityOf(adharno));
					wait30.until(ExpectedConditions.elementToBeClickable(adharno));
					String adharnop = adharno.getText();
				//	System.out.println(adharno.getText());
					
					if(adharno.getText().equals(adhar)) {
						System.out.println("yes");
					}
					
					WebDriverWait wait32 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement dateofbirth = wait32.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='DateofBirthPreview']")));
					wait32.until(ExpectedConditions.visibilityOf(dateofbirth));
					wait32.until(ExpectedConditions.elementToBeClickable(dateofbirth));
					String dobr = dateofbirth.getText();
				//	System.out.println(dateofbirth.getText());
					
					WebElement dateofbirthason = wait32.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='DateofBirthasonPreview']")));
					wait32.until(ExpectedConditions.visibilityOf(dateofbirthason));
					wait32.until(ExpectedConditions.elementToBeClickable(dateofbirthason));
					String dobas = dateofbirthason.getText();
				//	System.out.println(dateofbirthason.getText());
					
					WebDriverWait wait33 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement gender = wait33.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='genderPreview']")));
					wait33.until(ExpectedConditions.visibilityOf(gender));
					wait33.until(ExpectedConditions.elementToBeClickable(gender));
					String gend = gender.getText();
				//	System.out.println(gender.getText());
					
					//address

					WebDriverWait wait34 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement doorn = wait34.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='DoorPreview']")));
					wait34.until(ExpectedConditions.visibilityOf(doorn));
					wait34.until(ExpectedConditions.elementToBeClickable(doorn));
					String doornp = doorn.getText();
				//	System.out.println(doorn.getText());
					jss12.executeScript("arguments[0].scrollIntoView(true);", doorn);

					
					WebDriverWait wait341 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement streetper = wait341.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='StreetPreview']")));
					wait341.until(ExpectedConditions.visibilityOf(streetper));
					wait341.until(ExpectedConditions.elementToBeClickable(streetper));
					String streetpr = streetper.getText();
				//	System.out.println(streetper.getText());
					
					
					WebDriverWait wait36 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement landmarkpr = wait36.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='landmarkPreview']")));
					wait36.until(ExpectedConditions.visibilityOf(landmarkpr));
					wait36.until(ExpectedConditions.elementToBeClickable(landmarkpr));
					String land = landmarkpr.getText();
				//	System.out.println(landmarkpr.getText());
					
					WebDriverWait wait37 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement talukpr = wait341.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='talukPreview']")));
					wait37.until(ExpectedConditions.visibilityOf(talukpr));
					wait37.until(ExpectedConditions.elementToBeClickable(talukpr));
					String talukprr = talukpr.getText();
				//	System.out.println(talukpr.getText());
				    jss12.executeScript("arguments[0].scrollIntoView(true);", talukpr);

					
					WebDriverWait wait38 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement citypr = wait38.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='cityPreview']")));
					wait38.until(ExpectedConditions.visibilityOf(citypr));
					wait38.until(ExpectedConditions.elementToBeClickable(citypr));
					String cityy = citypr.getText();
				//	System.out.println(citypr.getText());
					
					
					
					WebDriverWait wait39 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement statep = wait39.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='statePreview']")));
					wait39.until(ExpectedConditions.visibilityOf(statep));
					wait39.until(ExpectedConditions.elementToBeClickable(statep));
					String statee = statep.getText();
				//	System.out.println(statep.getText());
					
					
					WebDriverWait wait40 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement districtp = wait38.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='districtPreview']")));
					wait38.until(ExpectedConditions.visibilityOf(districtp));
					wait38.until(ExpectedConditions.elementToBeClickable(districtp));
					String distr = districtp.getText();
				//	System.out.println(districtp.getText());
					 jss12.executeScript("arguments[0].scrollIntoView(true);", districtp);
					
					
					

					WebDriverWait wait43 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement pincodepr = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='pincodePreview']")));
					wait43.until(ExpectedConditions.visibilityOf(pincodepr));
					wait43.until(ExpectedConditions.elementToBeClickable(pincodepr));
					String pin = pincodepr.getText();
				//	System.out.println(pincodepr.getText());
					 jss12.executeScript("arguments[0].scrollIntoView(true);", pincodepr);

					
					WebDriverWait wait41 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement ndistrictp = wait41.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='nativeDistrictPreview']")));
					wait41.until(ExpectedConditions.visibilityOf(ndistrictp));
					wait41.until(ExpectedConditions.elementToBeClickable(ndistrictp));
					String nativee = ndistrictp.getText();
				//	System.out.println(ndistrictp.getText());
					
					
					WebDriverWait wait42 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement addsame = wait42.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='postaladdressPreview']")));
					wait42.until(ExpectedConditions.visibilityOf(addsame));
					wait42.until(ExpectedConditions.elementToBeClickable(addsame));
					String addres = addsame.getText();
				//	System.out.println(addsame.getText());
					
					
					//Permanent
					WebElement doornper = wait341.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='PerDoorNoPreview']")));
					wait341.until(ExpectedConditions.visibilityOf(doornper));
					wait341.until(ExpectedConditions.elementToBeClickable(doornper));
					String doo = doornper.getText();
				//	System.out.println(doornper.getText());
					jss12.executeScript("arguments[0].scrollIntoView(true);", doornper);

					WebElement streetper1 = wait341.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='PerStreetPreview']")));
					wait341.until(ExpectedConditions.visibilityOf(streetper1));
					wait341.until(ExpectedConditions.elementToBeClickable(streetper1));
					String streett = streetper1.getText();
				//	System.out.println(streetper1.getText());

					
					WebElement PerTalukPreview = wait36.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='PerTalukPreview']")));
					wait36.until(ExpectedConditions.visibilityOf(PerTalukPreview));
					wait36.until(ExpectedConditions.elementToBeClickable(PerTalukPreview));
				//	System.out.println(PerTalukPreview.getText());
					String talukk = PerTalukPreview.getText();
					

					WebElement PerCityPreview = wait341.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='PerCityPreview']")));
					wait37.until(ExpectedConditions.visibilityOf(PerCityPreview));
					wait37.until(ExpectedConditions.elementToBeClickable(PerCityPreview));
					String percity = PerCityPreview.getText();
				//	System.out.println(PerCityPreview.getText());
				    jss12.executeScript("arguments[0].scrollIntoView(true);", PerCityPreview);

					
					WebElement perstatePreview = wait38.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='perstatePreview']")));
					wait38.until(ExpectedConditions.visibilityOf(perstatePreview));
					wait38.until(ExpectedConditions.elementToBeClickable(perstatePreview));
					String perstate = perstatePreview.getText();
				//	System.out.println(perstatePreview.getText());
					
					
					WebElement PerotherDistrictPreview = wait39.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='PerotherDistrictPreview']")));
					wait39.until(ExpectedConditions.visibilityOf(PerotherDistrictPreview));
					wait39.until(ExpectedConditions.elementToBeClickable(PerotherDistrictPreview));
				//	System.out.println(PerotherDistrictPreview.getText());
					String otherdis = PerotherDistrictPreview.getText();
					
					
					WebElement PerPincodePreview = wait38.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='PerPincodePreview']")));
					wait38.until(ExpectedConditions.visibilityOf(PerPincodePreview));
					wait38.until(ExpectedConditions.elementToBeClickable(PerPincodePreview));
				//	System.out.println(PerPincodePreview.getText());
					String perpin = PerPincodePreview.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", PerPincodePreview);
					
				
					WebElement NearbyLandmarkPreview = wait41.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='NearbyLandmarkPreview']")));
					wait41.until(ExpectedConditions.visibilityOf(NearbyLandmarkPreview));
					wait41.until(ExpectedConditions.elementToBeClickable(NearbyLandmarkPreview));
				//	System.out.println(NearbyLandmarkPreview.getText());
					String nearlandmark = NearbyLandmarkPreview.getText();
					
					
					WebDriverWait wait44 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement castep = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='CategoryPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(castep));
					wait43.until(ExpectedConditions.elementToBeClickable(castep));
				//	System.out.println(castep.getText());
					String castee = castep.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", castep);
					
					
					WebDriverWait wait45 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement subcastep = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='SubcastePreview']")));
					wait43.until(ExpectedConditions.visibilityOf(subcastep));
					wait43.until(ExpectedConditions.elementToBeClickable(subcastep));
				//	System.out.println(subcastep.getText());
					String sub = subcastep.getText();
					
					
					WebDriverWait wait46 = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement dateofcastep = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='DateofSubcastePreview']")));
					wait43.until(ExpectedConditions.visibilityOf(dateofcastep));
					wait43.until(ExpectedConditions.elementToBeClickable(dateofcastep));
				//	System.out.println(dateofcastep.getText());
					String dateofsubcaste = dateofcastep.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", dateofcastep);
					
			
					WebElement govtemp = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='GovernmentEmployeePreview']")));
					wait43.until(ExpectedConditions.visibilityOf(govtemp));
					wait43.until(ExpectedConditions.elementToBeClickable(govtemp));
				//	System.out.println(govtemp.getText());
					String govermentemp = govtemp.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", govtemp);
					
					
					WebElement dojp = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='GovtDateofJoiningPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(dojp));
					wait43.until(ExpectedConditions.elementToBeClickable(dojp));
				//	System.out.println(dojp.getText());
					String dateofjion = dojp.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", dojp);

					
					WebElement govtdept = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='GovtDepartmentPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(govtdept));
					wait43.until(ExpectedConditions.elementToBeClickable(govtdept));
				//	System.out.println(govtdept.getText());
					String govermentdept = govtdept.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", govtdept);
					
					
					WebElement yearofserv = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='GovtYearsPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(yearofserv));
					wait43.until(ExpectedConditions.elementToBeClickable(yearofserv));
					String yearofser = yearofserv.getText();
			//		System.out.println(yearofserv.getText());
					
				WebElement monthofserv = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='GovtMonthPreview']")));
				wait43.until(ExpectedConditions.visibilityOf(monthofserv));
				wait43.until(ExpectedConditions.elementToBeClickable(monthofserv));
				String monthserv = monthofserv.getText();
			//	System.out.println(monthofserv.getText());

				WebElement dayofserv = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='GovtDaysPreview']")));
			    wait43.until(ExpectedConditions.visibilityOf(dayofserv));
			    wait43.until(ExpectedConditions.elementToBeClickable(dayofserv));
			    String dayser = dayofserv.getText();
			//	System.out.println(dayofserv.getText());
				
				
					
					WebElement desggovt = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='GovtDesignationPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(desggovt));
					wait43.until(ExpectedConditions.elementToBeClickable(desggovt));
			//		System.out.println(desggovt.getText());
					String dessgovt = desggovt.getText();
					
					
//					WebElement deptenq = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='DepartmentalEnquirPreview']")));
//					wait43.until(ExpectedConditions.visibilityOf(deptenq));
//					wait43.until(ExpectedConditions.elementToBeClickable(deptenq));
//				//	System.out.println(deptenq.getText());
//					String depen = deptenq.getText();
//					 jss12.executeScript("arguments[0].scrollIntoView(true);", deptenq);
//
//					
//					WebElement enquirydeat = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='DeptenqdetailsPreview']")));
//					wait43.until(ExpectedConditions.visibilityOf(enquirydeat));
//					wait43.until(ExpectedConditions.elementToBeClickable(enquirydeat));
//				//	System.out.println(enquirydeat.getText());
//					String enqdetails = enquirydeat.getText();
					// jss12.executeScript("arguments[0].scrollIntoView(true);", enquirydeat);
					
					WebElement criminalcase = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='CriminalCasesPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(criminalcase));
					wait43.until(ExpectedConditions.elementToBeClickable(criminalcase));
				//	System.out.println(criminalcase.getText());
					String crimcase = criminalcase.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", criminalcase);

				
					WebElement details = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='CriminalCasesdetailsPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(details));
					wait43.until(ExpectedConditions.elementToBeClickable(details));
			//		System.out.println(details.getText());
					String crimedetails = details.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", details);
					
					WebElement convcric = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='ConvictedinaCriminalPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(convcric));
					wait43.until(ExpectedConditions.elementToBeClickable(convcric));
				//	System.out.println(convcric.getText());
					String convcrime = convcric.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", convcric);
					
					
					WebElement ccdetails = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='ConvictedCriminalDetailsPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(ccdetails));
					wait43.until(ExpectedConditions.elementToBeClickable(ccdetails));
			//		System.out.println(ccdetails.getText());
					String crimedet = ccdetails.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", ccdetails);
					
					//sslc
					WebElement sslcpassed = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='PassedSSLCPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(sslcpassed));
					wait43.until(ExpectedConditions.elementToBeClickable(sslcpassed));
				//	System.out.println(sslcpassed.getText());
					String sslcpass = sslcpassed.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", sslcpassed);

					
					WebElement sslcboardp = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='BoardofSslcPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(sslcboardp));
					wait43.until(ExpectedConditions.elementToBeClickable(sslcboardp));
				//	System.out.println(sslcboardp.getText());
					String sslcboard = sslcboardp.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", sslcboardp);
					 
					 
                    if(board1==3) {
					 WebElement sslcboardoth = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='OtherSslcBoarPreview']")));
						wait43.until(ExpectedConditions.visibilityOf(sslcboardoth));
						wait43.until(ExpectedConditions.elementToBeClickable(sslcboardoth));
				//		System.out.println(sslcboardoth.getText());
						String otherboard = sslcboardoth.getText();
						 jss12.executeScript("arguments[0].scrollIntoView(true);", sslcboardp);
						  row1.createCell(50).setCellValue(otherboard);
                    }
					
					WebElement kannadalang = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='KannadaLanguagePreview']")));
					wait43.until(ExpectedConditions.visibilityOf(kannadalang));
					wait43.until(ExpectedConditions.elementToBeClickable(kannadalang));
				//	System.out.println(kannadalang.getText());
					String kannadalan = kannadalang.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", kannadalang);

					
					
					WebElement yearofpass = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='YearofPassingSSLCPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(yearofpass));
					wait43.until(ExpectedConditions.elementToBeClickable(yearofpass));
				//	System.out.println(yearofpass.getText());
					String yearofpassp = yearofpass.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", yearofpass);
					
					WebElement markorgrade = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='SSLCMarksorGradesPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(markorgrade));
					wait43.until(ExpectedConditions.elementToBeClickable(markorgrade));
				//	System.out.println(markorgrade.getText());
					String makorgra = markorgrade.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", markorgrade);

					
					WebElement maxmarkp = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='SSLCMaxMarksPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(maxmarkp));
					wait43.until(ExpectedConditions.elementToBeClickable(maxmarkp));
				//	System.out.println(maxmarkp.getText());
					String max = maxmarkp.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", maxmarkp);
					
					
					WebElement obatmark = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='SSLCMarksObtainedPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(obatmark));
					wait43.until(ExpectedConditions.elementToBeClickable(obatmark));
				//	System.out.println(obatmark.getText());
					String ob = obatmark.getText();
					// jss12.executeScript("arguments[0].scrollIntoView(true);", obatmark);
					
					WebElement persslc = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='SSLCPercentageObtainedPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(persslc));
					wait43.until(ExpectedConditions.elementToBeClickable(persslc));
				//	System.out.println(persslc.getText());
					String persslco = persslc.getText();
					 jss12.executeScript("arguments[0].scrollIntoView(true);", persslc);
					
					WebElement sslcreg = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='SSLCRegistrationNoPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(sslcreg));
					wait43.until(ExpectedConditions.elementToBeClickable(sslcreg));
				//	System.out.println(sslcreg.getText());
					String sslcregg = sslcreg.getText();
					// jss12.executeScript("arguments[0].scrollIntoView(true);", sslcreg);
					
					//puc
					WebElement passpuc = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='PassedPUCPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(passpuc));
					wait43.until(ExpectedConditions.elementToBeClickable(passpuc));
				//	System.out.println(passpuc.getText());
					String passpu = passpuc.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", passpuc);

					
					
					WebElement pucboard = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='PassedPucBoardPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(pucboard));
					wait43.until(ExpectedConditions.elementToBeClickable(pucboard));
				//	System.out.println(pucboard.getText());
					String puboardd = pucboard.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", pucboard);

					if(puboard1==3) {
					WebElement otherpuc = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='PucOtherBoardPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(otherpuc));
					wait43.until(ExpectedConditions.elementToBeClickable(otherpuc));
				//	System.out.println(otherpuc.getText());
					String otherpu = otherpuc.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", pucboard);
				     row1.createCell(62).setCellValue(otherpu);
					}

					WebElement yearpasspuc = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='YearofPassingPUCPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(yearpasspuc));
					wait43.until(ExpectedConditions.elementToBeClickable(yearpasspuc));
				//	System.out.println(yearpasspuc.getText());
					String yearpuc = yearpasspuc.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", yearpasspuc);

					
					
					WebElement markorgradepuc = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='MarksorGradePUCPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(markorgradepuc));
					wait43.until(ExpectedConditions.elementToBeClickable(markorgradepuc));
				//	System.out.println(markorgradepuc.getText());
					String pucmorg = markorgradepuc.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", markorgradepuc);

					
					
					WebElement gradepuc = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='GradesObtainedPUCPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(gradepuc));
					wait43.until(ExpectedConditions.elementToBeClickable(gradepuc));
				//	System.out.println(gradepuc.getText());
					String grADE = gradepuc.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", gradepuc);

					
					
					WebElement cgpapuc = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='PercentagePUCPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(cgpapuc));
					wait43.until(ExpectedConditions.elementToBeClickable(cgpapuc));
				//	System.out.println(cgpapuc.getText());
					String cgpapu = cgpapuc.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", cgpapuc);

					
					WebElement pucregp = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='PUCRegistrationsnoPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(pucregp));
					wait43.until(ExpectedConditions.elementToBeClickable(pucregp));
				//	System.out.println(pucregp.getText());
					String regpu = pucregp.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", pucregp);
				
			
					//deg
					WebElement DegHolderPreview = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='DegHolderPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(DegHolderPreview));
					wait43.until(ExpectedConditions.elementToBeClickable(DegHolderPreview));
				//	System.out.println(DegHolderPreview.getText());
					String degr = DegHolderPreview.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", DegHolderPreview);
					// Create a new row for each applicant
						//Row header = sheet2.createRow(0);
				     


					//qualification
					WebElement modaltypist = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='modaltypist']")));
					wait43.until(ExpectedConditions.visibilityOf(modaltypist));
					wait43.until(ExpectedConditions.elementToBeClickable(modaltypist));
				//	System.out.println(modaltypist.getText());
					String typ = modaltypist.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", modaltypist);

					
					
					WebElement Typsitqualimaxmarks = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='Typsitqualimaxmarks']")));
					wait43.until(ExpectedConditions.visibilityOf(Typsitqualimaxmarks));
					wait43.until(ExpectedConditions.elementToBeClickable(Typsitqualimaxmarks));
				//	System.out.println(Typsitqualimaxmarks.getText());
					String optyp = Typsitqualimaxmarks.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", Typsitqualimaxmarks);

					
					if (applyingpost != 2) {
					WebElement seniorshorthand = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='modalseniorshorthand']")));
					wait43.until(ExpectedConditions.visibilityOf(seniorshorthand));
					wait43.until(ExpectedConditions.elementToBeClickable(seniorshorthand));
				//	System.out.println(seniorshorthand.getText());
					String senr = seniorshorthand.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", seniorshorthand);

					
					WebElement seniorshorthandopt = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='Typsitqualiminimarks']")));
					wait43.until(ExpectedConditions.visibilityOf(seniorshorthandopt));
					wait43.until(ExpectedConditions.elementToBeClickable(seniorshorthandopt));
				//	System.out.println(seniorshorthandopt.getText());
					String senopt = seniorshorthandopt.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", seniorshorthandopt);
					
					row1.createCell(74).setCellValue(senr);
	                row1.createCell(75).setCellValue(senopt);
					}
					//documents
					WebElement IDCardPreview = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='IDCardSelectedPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(IDCardPreview));
					wait43.until(ExpectedConditions.elementToBeClickable(IDCardPreview));
				//	System.out.println(IDCardPreview.getText());
					String idcardp = IDCardPreview.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", IDCardPreview);

					WebElement IDCardNoPreview = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='SelectedIDCardNoPreview']")));
					wait43.until(ExpectedConditions.visibilityOf(IDCardNoPreview));
					wait43.until(ExpectedConditions.elementToBeClickable(IDCardNoPreview));
				//	System.out.println(IDCardNoPreview.getText());
					String idp = IDCardNoPreview.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", IDCardNoPreview);
					
					WebElement Identitymark01Preview = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='Identitymark01Preview']")));
					wait43.until(ExpectedConditions.visibilityOf(Identitymark01Preview));
					wait43.until(ExpectedConditions.elementToBeClickable(Identitymark01Preview));
				//	System.out.println(Identitymark01Preview.getText());
					String idmk = Identitymark01Preview.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", Identitymark01Preview);
					if(Identitymark01Preview.getText().equals(idmk1)) {
						//System.out.println("yes");
					}

					WebElement Identitymark02Preview = wait43.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='Identitymark02Preview']")));
					wait43.until(ExpectedConditions.visibilityOf(Identitymark02Preview));
					wait43.until(ExpectedConditions.elementToBeClickable(Identitymark02Preview));
				//	System.out.println(Identitymark02Preview.getText());
					String idmk02 = Identitymark02Preview.getText();
					jss12.executeScript("arguments[0].scrollIntoView(true);", Identitymark02Preview);
					
				    row1.createCell(1).setCellValue(appnamep);
                    row1.createCell(2).setCellValue(can);
                    row1.createCell(3).setCellValue(fathernamep);
                    row1.createCell(4).setCellValue(mothernamep);
                    row1.createCell(5).setCellValue(emailidp);
                    row1.createCell(6).setCellValue(mobilenop);
                    row1.createCell(7).setCellValue(adharnop);
                    row1.createCell(8).setCellValue(dobr);
                    row1.createCell(9).setCellValue(dobas);
                    row1.createCell(10).setCellValue(gend);
                    row1.createCell(11).setCellValue(doornp);
                    row1.createCell(12).setCellValue(streetpr);
                    row1.createCell(13).setCellValue(talukprr);
                    row1.createCell(14).setCellValue(cityy);
                    row1.createCell(15).setCellValue(statee);
                    row1.createCell(16).setCellValue(distr);
                    row1.createCell(18).setCellValue(pin);
                    row1.createCell(19).setCellValue(land);
                    row1.createCell(20).setCellValue(nativee);
                    row1.createCell(21).setCellValue(addres);
                    row1.createCell(22).setCellValue(doo);
                    row1.createCell(23).setCellValue(streett);
                    row1.createCell(24).setCellValue(talukk);
                    row1.createCell(25).setCellValue(percity);
                    row1.createCell(26).setCellValue(perstate);
                    row1.createCell(27).setCellValue(otherdis);
                    row1.createCell(29).setCellValue(perpin);
                    row1.createCell(30).setCellValue(nearlandmark);
                    row1.createCell(31).setCellValue(castee);
                    row1.createCell(32).setCellValue(sub);
                    row1.createCell(33).setCellValue(dateofsubcaste);
                    row1.createCell(34).setCellValue(govermentemp);
                    row1.createCell(35).setCellValue(dateofjion);
                    row1.createCell(36).setCellValue(govermentdept);
                    row1.createCell(38).setCellValue(yearofser);
                    row1.createCell(39).setCellValue(monthserv);
                    row1.createCell(40).setCellValue(dayser);
                    row1.createCell(41).setCellValue(dessgovt);
//                    row1.createCell(42).setCellValue(depen);
//                    row1.createCell(43).setCellValue(enqdetails);
                    row1.createCell(44).setCellValue(crimcase);
                    row1.createCell(45).setCellValue(crimedetails);
                    row1.createCell(46).setCellValue(convcrime);
                    row1.createCell(47).setCellValue(crimedet);
                    row1.createCell(48).setCellValue(sslcpass);
                    row1.createCell(49).setCellValue(sslcboard);
                  
                    row1.createCell(51).setCellValue(kannadalan);
                    row1.createCell(52).setCellValue(yearofpassp);
                    row1.createCell(53).setCellValue(makorgra);
                    row1.createCell(54).setCellValue(max);
                    row1.createCell(55).setCellValue(ob);
                    row1.createCell(56).setCellValue(persslco);
                    row1.createCell(59).setCellValue(sslcregg);
                    row1.createCell(60).setCellValue(passpu);
                    row1.createCell(61).setCellValue(puboardd);
              
                    row1.createCell(63).setCellValue(yearpuc);
                    row1.createCell(64).setCellValue(pucmorg);
                    row1.createCell(68).setCellValue(grADE);
                    row1.createCell(69).setCellValue(cgpapu);
                    row1.createCell(70).setCellValue(regpu);
                    row1.createCell(71).setCellValue(degr);
                    row1.createCell(72).setCellValue(typ);
                    row1.createCell(73).setCellValue(optyp);
                    if (applyingpost != 2) {
                   
                    }
                    row1.createCell(76).setCellValue(idcardp);
                    row1.createCell(77).setCellValue(idp);
                    row1.createCell(78).setCellValue(idmk);
                    row1.createCell(79).setCellValue(idmk02);
                   
                    driver.findElement(By.xpath("//button[text()='Submit']")).click();
					Thread.sleep(1500);
					Alert a = driver.switchTo().alert();
					a.getText();
					a.accept();
					Thread.sleep(1500);
				     // Switch to new window after clicking 'My Application'
			        switchToNewWindow(driver);

			        // Click on 'Forgot Application Number'
			        WebElement forgot = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a[contains(text(),'Forgot Application Number?')]")));
			        act.moveToElement(forgot).click().perform();

			        // Switch to new window after clicking 'Forgot Application Number'
			        switchToNewWindow(driver);

			        // Enter Aadhar number
			        WebElement adhar1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@class='form-control applNo']")));
			        adhar1.sendKeys(adhar);

			        // Enter Date of Birth
			        WebElement dateofbirth1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@class='form-control ldob dob flatpickr-input']")));
			        dateofbirth1.sendKeys(dob);

			        // Click 'Submit'
			        WebElement login = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a[text()='Submit ']")));
			        act.moveToElement(login).click().perform();

			        // Wait and scroll down after submission
			        Thread.sleep(1000);
			        js.executeScript("window.scrollBy(0,1000)");

			        // Wait for application number to appear
			        WebElement appno = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("(//tr[@class='odd' or @class='even']/td[1])[last()]")));
			        js.executeScript("window.scrollBy(0,2000)");
			        Thread.sleep(1000);
			        // Retrieve application number and print
			        String applno = appno.getText();
			      //  System.out.println(applno);
			        row1.createCell(80).setCellValue(applno);
			        // Close the current window
			        driver.findElement(By.linkText("Close")).click();
			        Thread.sleep(1000);

			        // Switch to main window after closing popup
			        switchToNewWindow(driver);

			        // Click on 'My Application' again
			        WebElement myaap1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.linkText("My Application")));
			        Thread.sleep(1000);
			        act.moveToElement(myaap1).click().perform();

			        // Switch to new window after clicking 'My Application'
			        switchToNewWindow(driver);

			        // Enter application number and date of birth
			        WebElement apno = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ApplicantModel_ApplicationNo")));
			        apno.sendKeys(applno);

			        WebElement dbo = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ApplicantModel_DateOfBirth")));
			        dbo.sendKeys(dob);

			        // Click 'Login'
			        WebElement log = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("login_submit")));
			        log.click();
			        Thread.sleep(4000);

			        // Switch to the correct window after login
			        switchToNewWindow(driver);

			        // Scroll down after login
			        js.executeScript("window.scrollBy(0,1000)");

			        // Wait for 'Download My Application' button to appear
			        WebElement download = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[text()='DOWNLOAD MY APPLICATION']")));
			        Thread.sleep(500);
			        // Click 'Download My Application'
			        act.moveToElement(download).click().perform();
			        Thread.sleep(1500);

					driver.findElement(By.linkText("Logout")).click();
					Thread.sleep(1500);
					driver.switchTo().window(Mainwindow);
					// Navigate to "New Application"
					driver.findElement(By.linkText("New Application")).click();
					Thread.sleep(2000);
					// Switch to the new window
					Set<String> allWindows1 = driver.getWindowHandles();
					for (String window : allWindows1) {
						driver.switchTo().window(window);
					}
					JavascriptExecutor js5 = (JavascriptExecutor) driver;
					js5.executeScript("window.scrollBy(0,1500);");
					Thread.sleep(2000);

					driver.findElement(By.xpath("//input[1]")).click();
					Thread.sleep(2000);

					driver.findElement(By.id("nextBtn")).click();
					allWindows = driver.getWindowHandles();
					for (String window : allWindows) {
						driver.switchTo().window(window);
					}
					

					js.executeScript("window.scrollBy(0,400);");
					Thread.sleep(1000);
					  // Save the workbook to the same Excel file
		            fileOut = new FileOutputStream("D://Automation_data//TestData (2).xlsx");
		            workbook.write(fileOut);///C://Users//pallavi//eclipse-workspace//project//TestData (2).xlsx
		            Reporter.log(i +" iteration succesfully completed");
		            System.out.println("ITERATION:");
		            System.out.println(i +" iteration succesfully completed ");
				}
				catch (Exception e) {
			        // Log the exception details
			        System.out.println("Error occurred in iteration " + i + ": " + e.getMessage());
			        e.printStackTrace();
			        
			        // Continue with the next iteration
			        Reporter.log("Failed :");
			        Reporter.log(i +" iteration is Skipping due to an error.");
			        
			        // Close the current application window
	                driver.findElement(By.linkText("Close")).click();
	                Thread.sleep(1000);
	                driver.switchTo().window(driver.getWindowHandle()); // Switch back to main window

	                // Navigate to "New Application" again
	                driver.findElement(By.linkText("New Application")).click();
	                Thread.sleep(2000);

	                // Switch to the new window for the next iteration
	                Set<String> allWindows1 = driver.getWindowHandles();
	                for (String window : allWindows1) {
	                    driver.switchTo().window(window);
	                }
	                JavascriptExecutor js5 = (JavascriptExecutor) driver;
					js5.executeScript("window.scrollBy(0,1500);");
					Thread.sleep(2000);

					driver.findElement(By.xpath("//input[1]")).click();
					Thread.sleep(2000);

					driver.findElement(By.id("nextBtn")).click();
					allWindows = driver.getWindowHandles();
					for (String window : allWindows) {
						driver.switchTo().window(window);
					}
					

					js.executeScript("window.scrollBy(0,400);");
					Thread.sleep(1000);

	                continue;  
				}
					
		
				}
        } }
        
        catch (IOException e) {
            e.printStackTrace(); // Log or handle the exception as necessary
        } finally {
            try {
                if (fileOut != null) {
                    fileOut.close();
                }
                if (fis != null) {
                    fis.close();
                }
                
            } catch (IOException e) {
                e.printStackTrace();
            }

            // Close workbook after operations are done
            // Note: Don't close the workbook until all operations are finished
            if (driver != null) {
              // driver.quit(); // Close the WebDriver session
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
                    SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy"); // Customize format
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
 private void switchToNewWindow(WebDriver driver) {
    Set<String> windowHandles = driver.getWindowHandles();
    for (String windowHandle : windowHandles) {
        driver.switchTo().window(windowHandle);
    }}
}

