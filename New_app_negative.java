package Steno;

import org.testng.annotations.Test;
import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Reporter;
import org.testng.annotations.Test;

public class New_app_negative {

    @Test
    public void sample() throws InterruptedException, IOException, AWTException {
        // System.setProperty("webdriver.chrome.driver", "./software/chromedriver.exe");
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
        Thread.sleep(2000);

        driver.findElement(By.xpath("//input[1]")).click();
        Thread.sleep(2000);

        driver.findElement(By.id("nextBtn")).click();
        allWindows = driver.getWindowHandles();
        for (String window : allWindows) {
            driver.switchTo().window(window);
        }
        Thread.sleep(2000);

        js.executeScript("window.scrollBy(0,400);");
        Thread.sleep(2000);

        // Open Excel file for reading and writing
        FileInputStream fis = null;
        FileOutputStream fileOut = null;
              
        try {
			// Reading Excel File
			FileInputStream fis1 = new FileInputStream("D://Automation_data//TestData (2).xlsx");// C://Users//pallavi//eclipse-workspace//project//Book5.xlsx
			XSSFWorkbook workbook = new XSSFWorkbook(fis1);//"C:\Users\pallavi\Desktop\TestData (2).xlsx"
			Sheet sheet = workbook.getSheetAt(1);

			// Select select = null;
			int rowCount = sheet.getPhysicalNumberOfRows();

			// Loop through rows in the Excel sheet
			// int rowCount = sheet.getPhysicalNumberOfRows();

			for (int i = 2; i <= 5; i++) { // Start from row 1 to skip header
				Row row = sheet.getRow(i);
				if (row == null) {
					
					System.out.println("Skipping empty row: " + i);
					continue;
				}
                if (row != null) {
                	
                        // fetching data from excel
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
		
						int applyingpost = (int) row.getCell(1).getNumericCellValue();
						String doornop = getCellValue(row.getCell(22));
						String streetp = getCellValue(row.getCell(23));
						
						String talukp = getCellValue(row.getCell(24));
						String cityp = getCellValue(row.getCell(25));
						int dropdownIndex5 = (int) row.getCell(27).getNumericCellValue();
						String pincodep = getCellValue(row.getCell(29));
						String landmarkp = getCellValue(row.getCell(30));
						int dropdownIndex6 = (int) row.getCell(26).getNumericCellValue();
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
						int typist1 = (int) row.getCell(73).getNumericCellValue();
						int kannada = (int) row.getCell(75).getNumericCellValue();
						int adharid = (int) row.getCell(76).getNumericCellValue();
						String idco = getCellValue(row.getCell(77));
						String idmk1 = getCellValue(row.getCell(78));
						String idmk2 = getCellValue(row.getCell(79));

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
						String othboard = getCellValue(row.getCell(62));
						String photo = getCellValue(row.getCell(80));
						String thumb1 = getCellValue(row.getCell(81));
						String idcard = getCellValue(row.getCell(82));

						
				//		System.out.println("Mobile Number from Excel: " + mob);
				//		System.out.println("DOB from Excel: " + dob);
					

				//		System.out.println("Application No: " + name);
				//		System.out.println("Father's Name: " + father);
						
				//	    System.out.println(applyingpost);
					    
					    WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(20));
						// Fill the form
					    
                         try {
						WebElement dropdown = driver.findElement(By.id("Applicant_PostUnitCode"));
						Select select1 = new Select(dropdown);
						select1.selectByIndex(applyingpost);

						//Applicant details
					    WebElement n = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_FullName")));
					    wait1.until(ExpectedConditions.visibilityOf(n));
					    wait1.until(ExpectedConditions.elementToBeClickable(n)); 
				        n.sendKeys(name);  
				        
					
						WebElement fathername = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_FatherName")));
					    wait1.until(ExpectedConditions.visibilityOf(fathername));
					    wait1.until(ExpectedConditions.elementToBeClickable(fathername)); 
					    fathername.sendKeys(father); 
						 WebElement errormessage = driver.findElement(By.xpath("//label[@id='Applicant_FullName-error']"));
					     String message = errormessage.getText(); 
					     
						
						
						WebElement mothername = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_MotherName")));
					    wait1.until(ExpectedConditions.visibilityOf(mothername));
					    wait1.until(ExpectedConditions.elementToBeClickable(mothername)); 
					    mothername.sendKeys(mother);  
						 WebElement errormessage1 = driver.findElement(By.xpath("//label[@id='Applicant_FatherName-error']"));
					     String message1 = errormessage1.getText();
					    
					     
						
						WebElement emailid = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EmailId")));
					    wait1.until(ExpectedConditions.visibilityOf(emailid));
					    wait1.until(ExpectedConditions.elementToBeClickable(emailid)); 
					    emailid.sendKeys(email);
					    
						 WebElement errormessage2 = driver.findElement(By.xpath("//label[@id='Applicant_MotherName-error']"));
					     String message2 = errormessage2.getText();
					    
					     
						WebElement mobile = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_MobileNo")));
					    wait1.until(ExpectedConditions.visibilityOf(mobile));
					    wait1.until(ExpectedConditions.elementToBeClickable(mobile)); 
					    mobile.sendKeys(mob);
						
						 WebElement errormessage3 = driver.findElement(By.xpath("//label[@id='Applicant_EmailId-error']"));
					     String message3 = errormessage3.getText();
					 //    System.out.println(message3);

						
						
						WebElement adharno = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_AadharNo")));
					    wait1.until(ExpectedConditions.visibilityOf(adharno));
					    wait1.until(ExpectedConditions.elementToBeClickable(adharno)); 
					    adharno.sendKeys(adhar);
					    
						 WebElement errormessage4 = driver.findElement(By.xpath("//label[@id='Applicant_MobileNo-error']"));
					     String message4 = errormessage4.getText();
					 //    System.out.println(message4);
						// Enter Date of Birth
						js.executeScript("window.scrollBy(0,200);");
						
						WebElement dobc = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_DateOfBirth")));
					    wait1.until(ExpectedConditions.visibilityOf(dobc));
					    wait1.until(ExpectedConditions.elementToBeClickable(dobc)); 
					    dobc.sendKeys(dob);
						Robot robot = new Robot();
						robot.keyPress(KeyEvent.VK_ENTER);
						robot.keyRelease(KeyEvent.VK_ENTER);
						 WebElement errormessage5 = driver.findElement(By.xpath("//label[@id='Applicant_AadharNo-error']"));
					     String message5 = errormessage5.getText();
					//     System.out.println(message5);
					     
					 	WebElement gen = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GenderCode")));
					    wait1.until(ExpectedConditions.visibilityOf(gen));
					    wait1.until(ExpectedConditions.elementToBeClickable(gen)); 
						Select ss1 = new Select(gen);
						Thread.sleep(1000);
						String gen1 = gen.getText();

						int dropdownIndex = (int) row.getCell(10).getNumericCellValue();
						
						// Select the dropdown option using the index
						System.out.println("Selecting option with index: " + dropdownIndex);
						ss1.selectByIndex(dropdownIndex);
						if(dobc.equals(""))
						{
						 WebElement errormessage6 = driver.findElement(By.xpath("//label[@id='Applicant_DateOfBirth-error']"));
					     String message6 = errormessage6.getText();
					//     System.out.println(message6);
						}
						Thread.sleep(500);
						
						
						// Applicant_ContactAddress
						
						WebElement dor = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_DoorNo")));
					    wait1.until(ExpectedConditions.visibilityOf(dor));
					    wait1.until(ExpectedConditions.elementToBeClickable(dor)); 
					    dor.sendKeys(doorno);
					    
						if(gen.equals(""))
						{
						WebElement errormessage7 = driver.findElement(By.xpath("//label[@id='Applicant_Reservation_GenderCode-error']"));
					     String message7 = errormessage7.getText();
					//     System.out.println(message7);
						}
						
						WebElement st = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_Street")));
					    wait1.until(ExpectedConditions.visibilityOf(st));
					    wait1.until(ExpectedConditions.elementToBeClickable(st)); 
					    st.sendKeys(street);
					    
						
						WebElement errormessage8 = driver.findElement(By.xpath("//label[@id='Applicant_ContactAddress_DoorNo-error']"));
					     String message8 = errormessage8.getText();
					//     System.out.println(message8);
					     
					     Actions act = new Actions(driver);
						
						js.executeScript("window.scrollBy(0,200)", "");
						
						WebElement t = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_Taluk")));
					    wait1.until(ExpectedConditions.visibilityOf(t));
					    wait1.until(ExpectedConditions.elementToBeClickable(t)); 
					    t.sendKeys(taluk);
						WebElement errormessage9 = driver.findElement(By.xpath("//label[@id='Applicant_ContactAddress_Street-error']"));
					     String message9 = errormessage9.getText();
					//     System.out.println(message9);
					     
					     
					
						WebElement ct = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_City")));
					    wait1.until(ExpectedConditions.visibilityOf(ct));
					    wait1.until(ExpectedConditions.elementToBeClickable(ct)); 
					    ct.sendKeys(city);
						
						WebElement errormessage11 = driver.findElement(By.xpath("//label[@id='Applicant_ContactAddress_Taluk-error']"));
					    String message11= errormessage11.getText();
					//    System.out.println(message11);
					     
						WebElement stt = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_UnionStateCode")));
					    wait1.until(ExpectedConditions.visibilityOf(stt));
					    wait1.until(ExpectedConditions.elementToBeClickable(stt)); 
						act.moveToElement(stt).click().perform();
						
						Select ss2 = new Select(stt);
						Thread.sleep(500);
						String state = stt.getText();
				//		WebElement errormessage12 = driver.findElement(By.xpath("//label[@id='Applicant_ContactAddress_City-error']"));
				//	    String message12= errormessage12.getText();
				//	    System.out.println(message12);
					     
					  Sheet sheet2 = workbook.getSheetAt(4);
				      Row row1 = sheet2.createRow(sheet2.getPhysicalNumberOfRows()); 
					  Robot r=new Robot();
						int dropdownIndex1 = (int) row.getCell(15).getNumericCellValue();

						// Add the condition to skip index 1
						if (dropdownIndex1 != 23) { // Only proceed if the index is not 1
						    ss2.selectByIndex(dropdownIndex1);  // Select the dropdown item by the given index
						 //   System.out.println("Selected index: " + dropdownIndex1);
						} else {
						   // System.out.println("Skipping index 1.");
						}

						// Scroll the window (optional)
						js.executeScript("window.scrollBy(0,300)", "");
						
						// Applicant_Reservation_CategoryCode

						
						WebElement dis = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_DistrictCode")));
					    wait1.until(ExpectedConditions.visibilityOf(dis));
					    wait1.until(ExpectedConditions.elementToBeClickable(dis)); 
					
						Select ss3 = new Select(dis);
						Thread.sleep(500);
						act.moveToElement(dis).click().perform();
						ss3.selectByIndex(dropdownIndex2);
						
						String district = dis.getText();
						if (dropdownIndex2 == 32) { // Only proceed if the index is not 1
						    ss3.selectByIndex(dropdownIndex2);  // Select the dropdown item by the given index
					//	    System.out.println("Selected index: " + dropdownIndex2);
						    WebDriverWait wait0 = new WebDriverWait(driver, Duration.ofSeconds(20));

							WebElement otherddis = wait0.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_OtherDistrictName")));

							wait0.until(ExpectedConditions.visibilityOf(otherddis));
							wait0.until(ExpectedConditions.elementToBeClickable(otherddis));
							otherddis.sendKeys(othdis);
							  r.keyPress(KeyEvent.VK_TAB);
							  r.keyRelease(KeyEvent.VK_TAB);
							WebElement errormessage13 = driver.findElement(By.xpath("//label[@id='Applicant_ContactAddress_OtherDistrictName-error']"));
						    String message13= errormessage13.getText();
					//	    System.out.println(message13);
						    row1.createCell(15).setCellValue(message13);
						} else {
						//	System.out.println("no");
						}
						  
						

						WebElement pin = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("pincodeInput")));
					    wait1.until(ExpectedConditions.visibilityOf(pin));
					    wait1.until(ExpectedConditions.elementToBeClickable(pin)); 
					    pin.sendKeys(pincode);	
					    r.keyPress(KeyEvent.VK_TAB);
					    r.keyRelease(KeyEvent.VK_TAB);
					    Thread.sleep(1500);
					     WebElement errormessage14 = driver.findElement(By.xpath("//label[@id='pincodeInput-error']"));
						 String message14= errormessage14.getText();
				//		 System.out.println(message14);
						 
						WebElement lan = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_Landmark")));
					    wait1.until(ExpectedConditions.visibilityOf(lan));
					    wait1.until(ExpectedConditions.elementToBeClickable(lan)); 
					    lan.sendKeys(landmark);
						  r.keyPress(KeyEvent.VK_TAB);
						  r.keyRelease(KeyEvent.VK_TAB);
						WebElement errormessage10 = driver.findElement(By.xpath("//label[@id='Applicant_ContactAddress_Landmark-error']"));
					    String message10= errormessage10.getText();
				//	    System.out.println(message10);
					    
						
						 
						WebElement ndis = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_NativeDistrict")));
					    wait1.until(ExpectedConditions.visibilityOf(ndis));
					    wait1.until(ExpectedConditions.elementToBeClickable(ndis)); 
					    ndis.sendKeys(Ndistrict);
						
						//for permanent address
						js.executeScript("window.scrollBy(0,100)", "");
						WebElement add = driver.findElement(By.xpath("//input[@id='Applicant_ContactAddress_IsPermanentAddressSame' and @value='False']"));
						Thread.sleep(1000);
						act.moveToElement(add).click().perform();
						add.click();
						
						js.executeScript("window.scrollBy(0,300)", "");
						WebElement errormessage15 = driver.findElement(By.xpath("//label[@id='Applicant_NativeDistrict-error']"));
					    String message15= errormessage15.getText();
					
					
						
						WebElement doorp = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PermanentAddress_DoorNo")));
					    wait1.until(ExpectedConditions.visibilityOf(doorp));
					    wait1.until(ExpectedConditions.elementToBeClickable(doorp)); 
					    doorp.sendKeys(doornop);
						driver.findElement(By.id("Applicant_PermanentAddress_Street")).sendKeys(streetp);
					
						WebElement errormessage16 = driver.findElement(By.xpath("//label[@id='Applicant_PermanentAddress_DoorNo-error']"));
					    String message16= errormessage16.getText();
					//    System.out.println(message16);

					
						WebElement tp = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PermanentAddress_Taluk")));
					    wait1.until(ExpectedConditions.visibilityOf(tp));
					    wait1.until(ExpectedConditions.elementToBeClickable(tp)); 
					    tp.sendKeys(talukp);
				
						WebElement errormessage17 = driver.findElement(By.xpath("//label[@id='Applicant_PermanentAddress_Street-error']"));
					    String message17= errormessage17.getText();
					//    System.out.println(message17);

					
						WebElement cp = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PermanentAddress_City")));
					    wait1.until(ExpectedConditions.visibilityOf(cp));
					    wait1.until(ExpectedConditions.elementToBeClickable(cp)); 
					    cp.sendKeys(cityp);
						WebElement errormessage18 = driver.findElement(By.xpath("//label[@id='Applicant_PermanentAddress_Taluk-error']"));
					    String message18= errormessage18.getText();
					//    System.out.println(message18);

						

						WebElement stat = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PermanentAddress_UnionStateCode")));
					    wait1.until(ExpectedConditions.visibilityOf(stat));
					    wait1.until(ExpectedConditions.elementToBeClickable(stat)); 
					    
						
						act.moveToElement(stat).click().perform();
				
						Select ss30 = new Select(stat);
						Thread.sleep(500);
						ss30.selectByIndex(dropdownIndex6);
						if (dropdownIndex6 != 23) { // Only proceed if the index is not 1
						    ss2.selectByIndex(dropdownIndex6);  // Select the dropdown item by the given index
					//	    System.out.println("Selected index: " + dropdownIndex6);
						} else {
					//	    System.out.println("Skipping index 1.");
						}
					//	WebElement errormessage19 = driver.findElement(By.xpath("//label[@id='Applicant_PermanentAddress_City-error']"));
					 //   String message19= errormessage19.getText();
					//    System.out.println(message19);
					    js.executeScript("window.scrollBy(0,200)", "");


						
						WebElement dist = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PermanentAddress_DistrictCode")));
					    wait1.until(ExpectedConditions.visibilityOf(dist));
					    wait1.until(ExpectedConditions.elementToBeClickable(dist)); 
					    Select ss31 = new Select(dist);
					
						ss31.selectByIndex(dropdownIndex5);
						if (dropdownIndex5 == 32) { // Only proceed if the index is not 1
						    ss31.selectByIndex(dropdownIndex5);  // Select the dropdown item by the given index
					//	    System.out.println("Selected index: " + dropdownIndex5);
							WebDriverWait wait0 = new WebDriverWait(driver, Duration.ofSeconds(20));

							WebElement otherddis = wait0.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PermanentAddress_OtherDistrictName")));

							wait0.until(ExpectedConditions.visibilityOf(otherddis));
							wait0.until(ExpectedConditions.elementToBeClickable(otherddis));
							otherddis.sendKeys(othdis);
						    r.keyPress(KeyEvent.VK_TAB);
						    r.keyRelease(KeyEvent.VK_TAB);
							WebElement errormessage20 = driver.findElement(By.xpath("//label[@id='Applicant_PermanentAddress_OtherDistrictName-error']"));
						    String message20= errormessage20.getText();
						//    System.out.println(message20);
		                    row1.createCell(28).setCellValue(message20);

						} else {
						
						}
		

						WebElement pinp = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PermanentAddress_Pincode")));
					    wait1.until(ExpectedConditions.visibilityOf(pinp));
					    wait1.until(ExpectedConditions.elementToBeClickable(pinp)); 
					    pinp.sendKeys(pincodep);
					    r.keyPress(KeyEvent.VK_TAB);
					    r.keyRelease(KeyEvent.VK_TAB);
					    Thread.sleep(1500);
					    WebElement errormessage21 = driver.findElement(By.xpath("//label[@id='Applicant_PermanentAddress_Pincode-error']"));
					    String message21= errormessage21.getText();
				//	    System.out.println(message21);
					    
						WebElement landp = wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_PermanentAddress_Landmark']")));
					    wait1.until(ExpectedConditions.visibilityOf(landp));
					    wait1.until(ExpectedConditions.elementToBeClickable(landp)); 
					    landp.sendKeys(landmarkp);
	                    r.keyPress(KeyEvent.VK_TAB);
					    r.keyRelease(KeyEvent.VK_TAB);
					    WebElement errormessage22= driver.findElement(By.xpath("//label[@id='Applicant_PermanentAddress_Landmark-error']"));
					    String message22= errormessage22.getText();
				//	    System.out.println(message22);
					    js.executeScript("window.scrollBy(0,200)", "");
					    
					    
					    // Applicant_Reservation_details
						
						WebElement cast = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_CategoryCode")));
					    wait1.until(ExpectedConditions.visibilityOf(cast));
					    wait1.until(ExpectedConditions.elementToBeClickable(cast)); 
						act.moveToElement(cast).click().perform();
						
						Select ss4 = new Select(cast);
						ss4.selectByIndex(caste);
						
					
						WebElement subcast = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_SubCaste")));
					    wait1.until(ExpectedConditions.visibilityOf(subcast));
					    wait1.until(ExpectedConditions.elementToBeClickable(subcast)); 
					    subcast.sendKeys(subcaste);
						r.keyPress(KeyEvent.VK_TAB);
					    r.keyRelease(KeyEvent.VK_TAB);
					    WebElement errormessage23= driver.findElement(By.xpath("//label[@id='Applicant_Reservation_SubCaste-error']"));
					    String message23= errormessage23.getText();
				
					
						WebElement cd = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_CategoryCertificateIssuedDate")));
					    wait1.until(ExpectedConditions.visibilityOf(cd));
					    wait1.until(ExpectedConditions.elementToBeClickable(cd)); 
					    cd.sendKeys(issue);
						r.keyPress(KeyEvent.VK_ENTER);
						r.keyRelease(KeyEvent.VK_ENTER);
			

						js.executeScript("window.scrollBy(0,200)", "");
						
						WebDriverWait wait55 = new WebDriverWait(driver, Duration.ofSeconds(20));

						WebElement cdd = wait55.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_AreYouAGovermentEmployee' and @value='True']")));
						wait55.until(ExpectedConditions.visibilityOf(cdd));
						wait55.until(ExpectedConditions.elementToBeClickable(cdd));
						cdd.click();

						WebElement gv = wait55.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_GovermentServiceDetail_JoiningDate']")));
						wait55.until(ExpectedConditions.visibilityOf(gv));
						act.moveToElement(gv).click().perform();
					    gv.sendKeys(govtjoin);
						r.keyPress(KeyEvent.VK_ENTER);
						r.keyRelease(KeyEvent.VK_ENTER);
					
						
						
						WebElement deptm = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GovermentServiceDetail_Department")));
					    wait1.until(ExpectedConditions.visibilityOf(deptm));
					    wait1.until(ExpectedConditions.elementToBeClickable(deptm)); 
					    deptm.sendKeys(dept);
						r.keyPress(KeyEvent.VK_TAB);
					    r.keyRelease(KeyEvent.VK_TAB);
					    WebElement errormessage24= driver.findElement(By.xpath("//label[@id='Applicant_Reservation_GovermentServiceDetail_Department-error']"));
					    String message24= errormessage24.getText();
					//    System.out.println(message24);
					    
					
						WebElement desg = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GovermentServiceDetail_Designation")));
					    wait1.until(ExpectedConditions.visibilityOf(desg));
					    wait1.until(ExpectedConditions.elementToBeClickable(desg)); 
					    desg.sendKeys(designation);
						r.keyPress(KeyEvent.VK_TAB);
					    r.keyRelease(KeyEvent.VK_TAB);
					    WebElement errormessage25= driver.findElement(By.xpath("//label[@id='Applicant_Reservation_GovermentServiceDetail_Designation-error']"));
					    String message25= errormessage25.getText();
					//    System.out.println(message25);
						js.executeScript("window.scrollBy(0,500)", "");// Applicant_Reservation_CategoryCode
						WebElement year = driver.findElement(By.id("Applicant_Reservation_GovermentServiceDetail_YearsInService"));
						year.click();
						Thread.sleep(500);
						Select ss5 = new Select(year);
						ss5.selectByIndex(year_index);
						WebElement month = driver.findElement(By.id("Applicant_Reservation_GovermentServiceDetail_MonthsInService"));
						month.click();
						Thread.sleep(500);
						Select ss6 = new Select(month);
						ss6.selectByIndex(month_index);
						WebElement day = driver.findElement(By.id("Applicant_Reservation_GovermentServiceDetail_DaysInService"));
						day.click();
						Thread.sleep(500);
						Select ss7 = new Select(day);
						ss7.selectByIndex(day_index);
						WebElement errormessage26= driver.findElement(By.xpath("//label[@id='Applicant_Reservation_GovermentServiceDetail_MonthsInService-error']"));
					    String message26= errormessage26.getText();
				//	    System.out.println(message26);
					    
					
						WebElement ab = wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_CriminalActivity_HasDepartmentEnquiry' and @value='True']")));
					    wait1.until(ExpectedConditions.visibilityOf(ab));
						act.moveToElement(ab).click().perform();

					
						WebElement cdq = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_CriminalActivity_DepartmentEnquiryDetail")));
					    wait1.until(ExpectedConditions.visibilityOf(cdq));
					    wait1.until(ExpectedConditions.elementToBeClickable(cdq)); 
					    cdq.sendKeys(dep_eq);
						r.keyPress(KeyEvent.VK_TAB);
					    r.keyRelease(KeyEvent.VK_TAB);
					    WebElement errormessage27= driver.findElement(By.xpath("//label[@id='Applicant_CriminalActivity_DepartmentEnquiryDetail-error']"));
					    String message27= errormessage27.getText();
				//	    System.out.println(message27);
						js.executeScript("window.scrollBy(0,400)");
					
				        WebElement crr = driver.findElement(By.xpath("//input[@id='Applicant_CriminalActivity_IsInvolvedInCriminalActivity' and @value='True']"));
						Thread.sleep(1000);
					    act.moveToElement(crr).click().perform();
					
						WebElement cde = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_CriminalActivity_CaseDetail")));
					    wait1.until(ExpectedConditions.visibilityOf(cde));
					    wait1.until(ExpectedConditions.elementToBeClickable(cde)); 
					    cde.sendKeys(crime_detail);
						js.executeScript("window.scrollBy(0,200)", "");
						
						r.keyPress(KeyEvent.VK_TAB);
					    r.keyRelease(KeyEvent.VK_TAB);
					    WebElement errormessage28= driver.findElement(By.xpath("//label[@id='Applicant_CriminalActivity_CaseDetail-error']"));
					    String message28= errormessage28.getText();
					//    System.out.println(message28);
					    
						
						 WebElement crd = driver.findElement(By.xpath("//input[@id='Applicant_CriminalActivity_IsConvictedInCriminalCase' and @value='True']"));
						Thread.sleep(1000);
						act.moveToElement(crd).click().perform();
							
						WebElement cdd1 = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_CriminalActivity_ConvictionDetail")));
					    wait1.until(ExpectedConditions.visibilityOf(cdd1));
					    wait1.until(ExpectedConditions.elementToBeClickable(cdd1)); 
					    cdd1.sendKeys(CA_detail);
						r.keyPress(KeyEvent.VK_TAB);
					    r.keyRelease(KeyEvent.VK_TAB);
					    WebElement errormessage29= driver.findElement(By.xpath("//label[@id='Applicant_CriminalActivity_ConvictionDetail-error']"));
					    String message29= errormessage29.getText();
				//	    System.out.println(message29);

					    
					    // for sslc
					 	
						WebElement sslc = driver.findElement(By.xpath("//input[@id='Applicant_EducationalQualification_IsSSLCHolder' and @value='True']"));
						Thread.sleep(1000);
 						act.moveToElement(sslc).click().perform();
			
						

						WebElement board = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_QualificationBoardCode")));

						wait1.until(ExpectedConditions.visibilityOf(board));
						wait1.until(ExpectedConditions.elementToBeClickable(board));

						// Scroll to the element
						JavascriptExecutor js1 = (JavascriptExecutor) driver;
						Select ss8 = new Select(board);
						ss8.selectByIndex(board1);
				//		System.out.println("Dropdown option selected!");
						// Ensure board1 is equal to 3 before proceeding
				//		System.out.println("board1 value: " + board1);

						if (board1==3) {
						  
						    ss8.selectByIndex(board1); 
						    System.out.println("Selected index: " + board1);

						  
						    WebElement sslcother = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_OtherBoardName")));
						    wait1.until(ExpectedConditions.visibilityOf(sslcother));
						    wait1.until(ExpectedConditions.elementToBeClickable(sslcother)); 
					        sslcother.sendKeys(othersslc);  
					        r.keyPress(KeyEvent.VK_TAB);
						    r.keyRelease(KeyEvent.VK_TAB);
							 WebElement errormessage30= driver.findElement(By.xpath("//label[@id='Applicant_EducationalQualification_SSLCQualification_OtherBoardName-error']"));
							 String message30= errormessage30.getText();
					//		 System.out.println(message30);
			                    row1.createCell(50).setCellValue(message30);

						} else {
						    System.out.println("Condition not met: board1 is not equal to 3.");
						}
						
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

						WebDriverWait wait3 = new WebDriverWait(driver, Duration.ofSeconds(20));

						WebElement year3 = wait3.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_YearOfPassing")));

						wait3.until(ExpectedConditions.visibilityOf(year3));
						wait3.until(ExpectedConditions.elementToBeClickable(year3));

						// Scroll to the element
						JavascriptExecutor jss12 = (JavascriptExecutor) driver;
					//	jss12.executeScript("arguments[0].scrollIntoView(true);", year3);

						Select ss11 = new Select(year3);
						ss11.selectByIndex(sslcyear);

						WebDriverWait wait4 = new WebDriverWait(driver, Duration.ofSeconds(20));

						WebElement mark = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_MarkType")));
					
						act.moveToElement(mark).click().perform();

						
					
						WebElement maxmark = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_Score_Maximum")));
					
						wait1.until(ExpectedConditions.visibilityOf(maxmark));
						wait1.until(ExpectedConditions.elementToBeClickable(maxmark));
						maxmark.clear();
						maxmark.sendKeys(m_mark);
						
						r.keyPress(KeyEvent.VK_TAB);
					    r.keyRelease(KeyEvent.VK_TAB);
						 WebElement errormessage31= driver.findElement(By.xpath("//label[@id='Applicant_EducationalQualification_SSLCQualification_Score_Maximum-error']"));
						 String message31= errormessage31.getText();
					//	 System.out.println(message31);
						WebDriverWait wait21 = new WebDriverWait(driver, Duration.ofSeconds(20));
						
						WebElement omark = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_Score_Obtained")));
				
						wait1.until(ExpectedConditions.visibilityOf(omark));
						wait1.until(ExpectedConditions.elementToBeClickable(omark));
						omark.clear();
						omark.sendKeys(ob_mark);

					
						
						r.keyPress(KeyEvent.VK_TAB);
					    r.keyRelease(KeyEvent.VK_TAB);
						 WebElement errormessage32= driver.findElement(By.xpath("//label[@id='Applicant_EducationalQualification_SSLCQualification_Score_Obtained-error']"));
						 String message32= errormessage32.getText();
				//		 System.out.println(message32);
						WebDriverWait wait5 = new WebDriverWait(driver, Duration.ofSeconds(20));
						
						WebElement per = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_ScoredPercentage")));
				
						wait1.until(ExpectedConditions.visibilityOf(per));
						wait1.until(ExpectedConditions.elementToBeClickable(per));
						act.moveToElement(per).click().perform();
				
						WebDriverWait wait51 = new WebDriverWait(driver, Duration.ofSeconds(20));

						WebElement regno = wait51.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@name='Applicant.EducationalQualification.SSLCQualification.RegistrationNo' and @id='Applicant_EducationalQualification_SSLCQualification_RegistrationNo']")));
						wait51.until(ExpectedConditions.visibilityOf(regno));
						wait51.until(ExpectedConditions.elementToBeClickable(regno));

						JavascriptExecutor js11 = (JavascriptExecutor) driver;
					//	js11.executeScript("arguments[0].scrollIntoView(true);", regno);
						// js11.executeScript("arguments[0].click();", regno); // Ensure it's focused

						regno.clear();
						regno.sendKeys(reg_no);
					//	WebElement errormessage33= driver.findElement(By.xpath("//label[@id='Applicant_EducationalQualification_SSLCQualification_ScoredPercentage-error']"));
					//	 String message33= errormessage33.getText();
					//	 System.out.println(message33);
						 
						 r.keyPress(KeyEvent.VK_TAB);
					     r.keyRelease(KeyEvent.VK_TAB);
						 WebElement errormessage34= driver.findElement(By.xpath("//label[@id='Applicant_EducationalQualification_SSLCQualification_RegistrationNo-error']"));
						 String message34= errormessage34.getText();
					//	 System.out.println(message34);
						 
						 
						 //for puc
						 
						 WebElement puc = driver.findElement(By.xpath("//input[@id='Applicant_EducationalQualification_IsPUCHolder' and @value='True']"));
						 Thread.sleep(1000);
						 act.moveToElement(puc).click().perform();
						
							WebDriverWait wait11 = new WebDriverWait(driver, Duration.ofSeconds(20));
							
							WebElement puboard = wait11.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_QualificationBoardCode")));

							wait11.until(ExpectedConditions.visibilityOf(puboard));
							wait11.until(ExpectedConditions.elementToBeClickable(puboard));

							// Scroll to the element
						//	js.executeScript("arguments[0].scrollIntoView(true);", puboard);

							Select ss14 = new Select(puboard);
							ss14.selectByIndex(puboard1);
					//		System.out.println("Dropdown option selected!");

							if (puboard1==3) {
							   
							    ss14.selectByIndex(puboard1);  // Select the dropdown item by the given index
						//	    System.out.println("Selected index: " + puboard1);
							 
							    WebElement puother = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_OtherBoardName")));
							    wait1.until(ExpectedConditions.visibilityOf(puother));
							    wait1.until(ExpectedConditions.elementToBeClickable(puother));
							    puother.sendKeys(othboard);
							    r.keyPress(KeyEvent.VK_TAB);
							     r.keyRelease(KeyEvent.VK_TAB);
								 WebElement errormessage35= driver.findElement(By.xpath("//label[@id='Applicant_EducationalQualification_PUCQualification_OtherBoardName-error']"));
								 String message35= errormessage35.getText();
						//		 System.out.println(message35);
				                    row1.createCell(62).setCellValue(message35);

							} else {
							   
							    // Log to make sure the field is ready
							    System.out.println("Field is ready: " );
							    
							   
							}
							   
							WebDriverWait wait31 = new WebDriverWait(driver, Duration.ofSeconds(20));

							WebElement puyear3 = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_YearOfPassing")));

							wait31.until(ExpectedConditions.visibilityOf(puyear3));
							wait31.until(ExpectedConditions.elementToBeClickable(puyear3));

							// Scroll to the element
						//	jss12.executeScript("arguments[0].scrollIntoView(true);", puyear3);

							Select ss111 = new Select(puyear3);
							ss111.selectByIndex(pucyear);

							WebElement grade = wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_PUCQualification_MarkType' and @value='G']")));
							act.moveToElement(grade).click().perform();

							WebElement gr = wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_PUCQualification_Grade_Grade' and @name='Applicant.EducationalQualification.PUCQualification.Grade.Grade']")));
							wait1.until(ExpectedConditions.visibilityOf(gr));
							wait1.until(ExpectedConditions.elementToBeClickable(gr));
							gr.clear();
							gr.sendKeys(pugrade);
							 r.keyPress(KeyEvent.VK_TAB);
						     r.keyRelease(KeyEvent.VK_TAB);
					//	 WebElement errormessage36= driver.findElement(By.xpath("//label[@id='Applicant_EducationalQualification_PUCQualification_Grade_Grade-error']"));
						// String message36= errormessage36.getText();
					//	 System.out.println(message36);
					//	   
							
							WebElement per1 = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_ScorePercentage")));
							wait1.until(ExpectedConditions.visibilityOf(per1));
							wait1.until(ExpectedConditions.elementToBeClickable(per1));
							per1.clear();
							per1.sendKeys(cgpa);
							 r.keyPress(KeyEvent.VK_TAB);
						     r.keyRelease(KeyEvent.VK_TAB);
							 WebElement errormessage37= driver.findElement(By.xpath("//label[@id='Applicant_EducationalQualification_PUCQualification_ScorePercentage-error']"));
							 String message37= errormessage37.getText();
						//	 System.out.println(message37);
						   
							

							WebElement pureg = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_RegistrationNo")));
							wait1.until(ExpectedConditions.visibilityOf(pureg));
							wait1.until(ExpectedConditions.elementToBeClickable(pureg));
							pureg.clear();
							pureg.sendKeys(pucreg);
							 r.keyPress(KeyEvent.VK_TAB);
						     r.keyRelease(KeyEvent.VK_TAB);
							 WebElement errormessage38= driver.findElement(By.xpath("//label[@id='Applicant_EducationalQualification_PUCQualification_RegistrationNo-error']"));
							 String message38= errormessage38.getText();
				//			 System.out.println(message38);
							
							

							// for degree

							WebElement deg = driver.findElement(By.xpath("//input[@id='Applicant_EducationalQualification_IsDegreeHolder' and @value='True']"));
						   Thread.sleep(1000);
						     act.moveToElement(deg).click().perform();
							
							
							// for qualification

							WebElement senior = driver.findElement(By.xpath("//input[@id='Applicant_TypistAssistant_IsPassedInQualifyExam' and @value='True']"));
							Thread.sleep(1000);
							act.moveToElement(senior).click().perform();
							

							WebDriverWait wait6 = new WebDriverWait(driver, Duration.ofSeconds(20));
							WebElement opt11 = wait6.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_TypistAssistant_QualificationCode")));

							wait6.until(ExpectedConditions.visibilityOf(opt11));
							wait6.until(ExpectedConditions.elementToBeClickable(opt11));

						

							Select ss15 = new Select(opt11);
							ss15.selectByIndex(typist1);
							String selectedOption = ss15.getFirstSelectedOption().getText();

					

							if (applyingpost != 2) {
								WebElement typist = driver.findElement(By.xpath("//input[@id='Applicant_StenographerAssistant_IsPassedInQualifyExam' and @value='True']"));
								Thread.sleep(1000);

								act.moveToElement(typist).click().perform();
						

								WebDriverWait wait61 = new WebDriverWait(driver, Duration.ofSeconds(20));
								WebElement topt1 = wait61.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_StenographerAssistant_QualificationCode")));

								wait61.until(ExpectedConditions.visibilityOf(topt1));
								wait61.until(ExpectedConditions.elementToBeClickable(topt1));

								// Scroll to the dropdown element
								jss12.executeScript("arguments[0].scrollIntoView(true);", topt1);


								Select ss16 = new Select(topt1);
								ss16.selectByIndex(kannada);
							}
							
							
							// for document upload
							WebDriverWait wait17 = new WebDriverWait(driver, Duration.ofSeconds(20));
							WebElement dopt1 = wait17.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_IdentityCardTypeCode")));

							wait17.until(ExpectedConditions.visibilityOf(dopt1));
							wait17.until(ExpectedConditions.elementToBeClickable(dopt1));

							// Select the dropdown item by index (assuming adharid is 2)
							Select ss17 = new Select(dopt1);
							ss17.selectByIndex(adharid);
							Thread.sleep(500);  // Wait a moment to allow any dynamic behavior

							if (adharid != 1) {
							    // Only proceed if the index is not 1
							    ss17.selectByIndex(adharid);  // Select the dropdown item by the given index
						//	    System.out.println("Selected index: " + adharid);
							  //  / Wait for the text field to be present and interactable
							    WebElement idcod = wait17.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_UploadedIDNo")));
							    wait17.until(ExpectedConditions.visibilityOf(idcod));
							    wait17.until(ExpectedConditions.elementToBeClickable(idcod));
							    Thread.sleep(500); 
							    idcod.sendKeys(idco);
							    // Log to make sure the field is ready
						//	    System.out.println("Field is ready: " + idcod.isEnabled());
							    
							} else {
							   
							    // Log to make sure the field is ready
						//	    System.out.println("Field is ready: " );
							    
							   
							}
							 r.keyPress(KeyEvent.VK_TAB);
						     r.keyRelease(KeyEvent.VK_TAB);
							 WebElement errormessage39= driver.findElement(By.xpath("//label[@id='Applicant_UploadedIDNo-error']"));
							 String message39= errormessage39.getText();
					//		 System.out.println(message39);
						
							WebDriverWait wait20 = new WebDriverWait(driver, Duration.ofSeconds(20));

							WebDriverWait wait18 = new WebDriverWait(driver, Duration.ofSeconds(20));

							WebElement idtm = wait18.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_IdentificationMark_01")));
							wait18.until(ExpectedConditions.visibilityOf(idtm));
							wait18.until(ExpectedConditions.elementToBeClickable(idtm));
							idtm.clear();
							idtm.sendKeys(idmk1);
							 r.keyPress(KeyEvent.VK_TAB);
						     r.keyRelease(KeyEvent.VK_TAB);
							 WebElement errormessage40= driver.findElement(By.xpath("//label[@id='Applicant_IdentificationMark_01-error']"));
							 String message40= errormessage40.getText();
						//	 System.out.println(message40);
							
							WebDriverWait wait19 = new WebDriverWait(driver, Duration.ofSeconds(20));

							WebElement idtm2 = wait19.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_IdentificationMark_02")));
							wait19.until(ExpectedConditions.visibilityOf(idtm2));
							wait19.until(ExpectedConditions.elementToBeClickable(idtm2));
							idtm2.clear();
							idtm2.sendKeys(idmk2);
							 r.keyPress(KeyEvent.VK_TAB);
						     r.keyRelease(KeyEvent.VK_TAB);
							 WebElement errormessage41= driver.findElement(By.xpath("//label[@id='Applicant_IdentificationMark_02-error']"));
							 String message41= errormessage41.getText();
						//	 System.out.println(message41);
							if (applyingpost != 1) {
								js.executeScript("window.scrollBy(0,400)", "");
								Thread.sleep(4000);
								
							}
							
							WebElement file = driver.findElement(By.name("Applicant.Photo"));
							file.sendKeys(photo);
							 r.keyPress(KeyEvent.VK_TAB);
						     r.keyRelease(KeyEvent.VK_TAB);
							 WebElement errormessage42= driver.findElement(By.xpath("//label[contains(text(),'Image size should not exceed 400 kb.')]"));
							 String message42= errormessage42.getText();
						//	 System.out.println(message42);
							
							WebElement thumb = driver.findElement(By.name("Applicant.Thumb"));
							thumb.sendKeys(thumb1);
							 r.keyPress(KeyEvent.VK_TAB);
						     r.keyRelease(KeyEvent.VK_TAB);
						  
						     WebElement id = driver.findElement(By.name("Applicant.IdentityCard"));
							id.sendKeys(idcard);
						
							 js.executeScript("window.scrollBy(0,200)", "");
								WebElement pr = driver.findElement(By.id("preview-btn"));
								act.moveToElement(pr).click().perform();
								r.keyPress(KeyEvent.VK_ENTER);
								r.keyRelease(KeyEvent.VK_ENTER);
						
								
								
					WebElement errormessage43= driver.findElement(By.xpath("//label[contains(text(),'Image size should not exceed 150 kb.')]"));
					   String message43= errormessage43.getText();
					//   System.out.println(message43);
					   
					   WebElement errormessage44= driver.findElement(By.xpath("//label[contains(text(),'Image size should not exceed 400 kb.')]"));
					   String message44= errormessage44.getText();
				//	   System.out.println(message44);

					driver.findElement(By.xpath("//a[text()='Close']")).click();
				
				//	Alert a = driver.switchTo().alert();
				//	a.getText();
				//	a.accept();
				//	Thread.sleep(3000);
				//	driver.findElement(By.linkText("Close")).click();
				//	Thread.sleep(3000);
					driver.switchTo().window(Mainwindow);
					// Navigate to "New Application"
					driver.findElement(By.linkText("New Application")).click();
					Thread.sleep(1000);
					// Switch to the new window
					Set<String> allWindows1 = driver.getWindowHandles();
					for (String window : allWindows1) {
						driver.switchTo().window(window);
					}
					JavascriptExecutor js5 = (JavascriptExecutor) driver;
					js5.executeScript("window.scrollBy(0,1500);");
					Thread.sleep(1000);

					driver.findElement(By.xpath("//input[1]")).click();
					Thread.sleep(1000);

					driver.findElement(By.id("nextBtn")).click();
					allWindows = driver.getWindowHandles();
					for (String window : allWindows) {
						driver.switchTo().window(window);
					}
					Thread.sleep(1000);

					js.executeScript("window.scrollBy(0,400);");
					Thread.sleep(500);
	               

				
			     
             
                  //  row1.createCell(1).setCellValue(appnamep);
                    row1.createCell(2).setCellValue(message);
                    row1.createCell(3).setCellValue(message1);
                    row1.createCell(4).setCellValue(message2);
                    row1.createCell(5).setCellValue(message3);
                    row1.createCell(6).setCellValue(message4);
                    row1.createCell(7).setCellValue(message5);
                   // row1.createCell(8).setCellValue(message6);
                   // row1.createCell(9).setCellValue(dobas);
                   // row1.createCell(10).setCellValue(gend);
                    row1.createCell(11).setCellValue(message8);
                    row1.createCell(12).setCellValue(message9);
                    row1.createCell(13).setCellValue(message11);
             //       row1.createCell(14).setCellValue(message12);
                 
                   // row1.createCell(16).setCellValue(statee);
                   // row1.createCell(17).setCellValue(distr);
                    row1.createCell(18).setCellValue(message14);
                    row1.createCell(19).setCellValue(message10);
                    row1.createCell(20).setCellValue(message15);
                //    row1.createCell(21).setCellValue(addres);
                    row1.createCell(22).setCellValue(message16);
                    row1.createCell(23).setCellValue(message17);
                    row1.createCell(24).setCellValue(message18);
               //     row1.createCell(25).setCellValue(message19);
              //      row1.createCell(26).setCellValue(perstate);
                 //   row1.createCell(27).setCellValue(otherdis);
                    row1.createCell(29).setCellValue(message21);
                    row1.createCell(30).setCellValue(message22);
               //     row1.createCell(31).setCellValue(castee);
                    row1.createCell(32).setCellValue(message23);
                //    row1.createCell(33).setCellValue(dateofsubcaste);
               //     row1.createCell(34).setCellValue(govermentemp);
                //    row1.createCell(35).setCellValue(dateofjion);
                    row1.createCell(36).setCellValue(message24);
                //   row1.createCell(38).setCellValue(yearofser);
                 //   row1.createCell(39).setCellValue(monthserv);
                    row1.createCell(40).setCellValue(message26);
                    row1.createCell(41).setCellValue(message25);
                  //  row1.createCell(42).setCellValue(depen);
                    row1.createCell(43).setCellValue(message27);
                  //  row1.createCell(44).setCellValue(crimcase);
                    row1.createCell(45).setCellValue(message28);
                  //  row1.createCell(46).setCellValue(convcrime);
                    row1.createCell(47).setCellValue(message29);
               //     row1.createCell(48).setCellValue(sslcpass);
                //    row1.createCell(49).setCellValue(sslcboard);
                //    row1.createCell(51).setCellValue(kannadalan);
              //      row1.createCell(52).setCellValue(yearofpassp);
              //      row1.createCell(53).setCellValue(makorgra);
                    row1.createCell(54).setCellValue(message31);
                    row1.createCell(55).setCellValue(message32);
           //         row1.createCell(56).setCellValue(message33);
                    row1.createCell(59).setCellValue(message34);
                 //   row1.createCell(60).setCellValue(passpu);
                 //   row1.createCell(61).setCellValue(puboardd);
                 //   row1.createCell(63).setCellValue(yearpuc);
                 //   row1.createCell(64).setCellValue(pucmorg);
                  // row1.createCell(68).setCellValue(message36);
                    row1.createCell(69).setCellValue(message37);
                    row1.createCell(70).setCellValue(message38);
                 //   row1.createCell(71).setCellValue(degr);
                 //  row1.createCell(72).setCellValue(typ);
                 //   row1.createCell(73).setCellValue(optyp);
                    if (applyingpost != 2) {
                   
                    }
                  //  row1.createCell(76).setCellValue(idcardp);
                    row1.createCell(77).setCellValue(message39);
                    row1.createCell(78).setCellValue(message40);
                    row1.createCell(79).setCellValue(message41);
                    row1.createCell(80).setCellValue(message42);
                    row1.createCell(81).setCellValue(message43);
                    row1.createCell(82).setCellValue(message42);


            // Save the workbook to the same Excel file
            fileOut = new FileOutputStream("D://Automation_data//TestData (2).xlsx");
            workbook.write(fileOut);  // Write data to the file
            Reporter.log(i +" iteration succesfully completed");
            System.out.println(i + " iteration succesfully completed");
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
        } }}catch (IOException e) {
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
}

