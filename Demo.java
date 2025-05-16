package psi_sports;

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
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class Demo {
 
	@Test
	public void sample() throws InterruptedException, AWTException
	{
		ChromeDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("http://172.10.1.159:9033");
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
		wait.until(ExpectedConditions.presenceOfElementLocated(By.linkText("New Application"))).click();
		Set<String> window = driver.getWindowHandles();
		for (String allwin : window) {
			driver.switchTo().window(allwin);	
		}
		JavascriptExecutor jss = (JavascriptExecutor) driver;
		jss.executeScript("window.scrollBy(0,1900)", "");
	  	Thread.sleep(500);
		wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[1]"))).click();
		
		wait.until(ExpectedConditions.presenceOfElementLocated(By.id("nextBtn"))).click();
		Set<String> windows = driver.getWindowHandles();
		for (String allwind : windows) {
			driver.switchTo().window(allwind);
		}
		
		jss.executeScript("window.scrollBy(0,100)", "");
		Actions act = new Actions(driver);
		Robot r = new Robot();
		 FileInputStream fis = null;
		    FileOutputStream fileOut = null;

		    try {
				// Reading Excel File
				FileInputStream fis1 = new FileInputStream("D://Automation_data//PSI_Sports.xlsx");//"D:\steno\TestDataAPC.xlsx"
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
        
        	int applyingpost = (int) row.getCell(0).getNumericCellValue();
        	int Priority = (int) row.getCell(1).getNumericCellValue();
        	int Priority1 = (int) row.getCell(2).getNumericCellValue();
        	int Priority2 = (int) row.getCell(3).getNumericCellValue();
        	int Priority3 = (int) row.getCell(4).getNumericCellValue();
        	int Priority4 = (int) row.getCell(5).getNumericCellValue();
        	int Priority5 = (int) row.getCell(6).getNumericCellValue();
        	int Priority6 = (int) row.getCell(7).getNumericCellValue();
        	int Priority7 = (int) row.getCell(8).getNumericCellValue();
        	int Priority8 = (int) row.getCell(9).getNumericCellValue();
        	int Priority9 = (int) row.getCell(10).getNumericCellValue();
        	
        	String Applicant_FullNameE = getCellValue(row.	getCell(11));
			String Fathername = getCellValue(row.getCell(12));
			String Applicant_MotherName = getCellValue(row.getCell(13));
			String Applicant_EmailId = getCellValue(row.getCell(14));
			String Applicant_MobileNo = getCellValue(row.getCell(15));
			String Applicant_AadharNo = getCellValue(row.getCell(16));
			String date = getCellValue(row.getCell(17));
			int GenderCode =  (int) row.getCell(19).getNumericCellValue();
			String PhysicalDetail_Height = getCellValue(row.getCell(20));
			String PhysicalDetail_Weight = getCellValue(row.getCell(21));
			
			String DoorNo = getCellValue(row.getCell(22));
			String ContactAddress_Street = getCellValue(row.getCell(23));
			String ContactAddress_Taluk = getCellValue(row.getCell(24));
			String ContactAddress_City = getCellValue(row.getCell(25));
			int UnionStateCode =  (int) row.getCell(26).getNumericCellValue();
			int DistrictCode =  (int) row.getCell(27).getNumericCellValue();
			String Address_Pincode = getCellValue(row.getCell(29));
			String Address_Landmark = getCellValue(row.getCell(30));
			String NativeDistrict = getCellValue(row.getCell(31));
			
			String PermanentDoorNo = getCellValue(row.getCell(33));
			String PermanentAddress_Street = getCellValue(row.getCell(34));
			String PermanentAddress_Taluk = getCellValue(row.getCell(35));
			String PermanentAddress_City = getCellValue(row.getCell(36));
			int PermanentUnionStateCode =  (int) row.getCell(37).getNumericCellValue();
			int PermanentDistrictCode =  (int) row.getCell(38).getNumericCellValue();
			String Permanent_Pincode = getCellValue(row.getCell(40));
			String Permanent_Landmark = getCellValue(row.getCell(41));
			
			int CategoryCode =  (int) row.getCell(42).getNumericCellValue();
			String Reservation_SubCaste = getCellValue(row.getCell(43));
			String CategoryCertificateIssuedDate = getCellValue(row.getCell(44));
			int Sportsachivement =  (int) row.getCell(45).getNumericCellValue();
			int game =  (int) row.getCell(46).getNumericCellValue();
			String gamedetails = getCellValue(row.getCell(47));
			String KKCertificateIssuedDate = getCellValue(row.getCell(53));
			String joining_date = getCellValue(row.getCell(55));
			int YearsInService_govt =  (int) row.getCell(57).getNumericCellValue();
			int MonthsInService_govt =  (int) row.getCell(58).getNumericCellValue();
			int DaysInService_govt =  (int) row.getCell(59).getNumericCellValue();
			String GovermentServiceDetail_Department = getCellValue(row.getCell(60));
			String GovermentServiceDetail_Designation = getCellValue(row.getCell(61));
			String DepartmentEnquiryDetail = getCellValue(row.getCell(63));
			int KSP =  (int) row.getCell(65).getNumericCellValue();
			String applicantInService_Designation = getCellValue(row.getCell(66));
			String joining_date_service = getCellValue(row.getCell(67));
			
			
			
			
		WebElement ApplyingType = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplyingTypeCode")));
		Select s=new Select(ApplyingType);
		s.selectByIndex(applyingpost);
		
		WebElement prioriti = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantPostUnitsPriority_0__PostUnitCode")));
		Select p=new Select(prioriti);
		p.selectByIndex(Priority);
		Thread.sleep(500);
		WebElement prioriti1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantPostUnitsPriority_1__PostUnitCode")));
		Select p1=new Select(prioriti1);
		p1.selectByIndex(Priority1);
		Thread.sleep(500);
		WebElement prioriti2 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantPostUnitsPriority_2__PostUnitCode")));
		Select p2=new Select(prioriti2);
		p2.selectByIndex(Priority2);
		Thread.sleep(500);
		WebElement prioriti3 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantPostUnitsPriority_3__PostUnitCode")));
		Select p3=new Select(prioriti3);
		p3.selectByIndex(Priority3);
		
		WebElement prioriti4 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantPostUnitsPriority_4__PostUnitCode")));
		Select p4=new Select(prioriti4);
		p4.selectByIndex(Priority4);
		
		WebElement prioriti5 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantPostUnitsPriority_5__PostUnitCode")));
		Select p5=new Select(prioriti5);
		p5.selectByIndex(Priority5);
		
		WebElement prioriti6 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantPostUnitsPriority_6__PostUnitCode")));
		Select p6=new Select(prioriti6);
		p6.selectByIndex(Priority6);
		
		WebElement prioriti7 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantPostUnitsPriority_7__PostUnitCode")));
		Select p7=new Select(prioriti7);
		p7.selectByIndex(Priority7);
		
		WebElement prioriti8 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantPostUnitsPriority_8__PostUnitCode")));
		Select p8=new Select(prioriti8);
		p8.selectByIndex(Priority8);
		
		WebElement prioriti9 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantPostUnitsPriority_9__PostUnitCode")));
		Select p9=new Select(prioriti9);
		p9.selectByIndex(Priority9);
		Thread.sleep(1000);
		
		//For Personal Details 
		
		WebElement Applicant_FullName = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_FullName")));
		Applicant_FullName.sendKeys(Applicant_FullNameE);
		
		WebElement Father = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_FatherName")));
		Father.sendKeys(Fathername);
		jss.executeScript("window.scrollBy(0,400)", "");
		
		WebElement Mother = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_MotherName")));
		Mother.sendKeys(Applicant_MotherName);
		
		WebElement Email = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EmailId")));
		Email.sendKeys(Applicant_EmailId);
		
		WebElement Mob = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_MobileNo")));
		Mob.sendKeys(Applicant_MobileNo);
		
		WebElement adhar = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_AadharNo")));
		adhar.sendKeys(Applicant_AadharNo);
		

        Thread.sleep(500);
        WebElement datePicker = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("(//input[@class='form-control form-control input' and @type='text' and @readonly='readonly'])[1]")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", datePicker);
        Dob(date, driver, "(//input[@class='numInput cur-year'])[1]", "(//select[contains(@class,'flatpickr-monthDropdown-months')])[1]");

//        
		
		WebElement Gender = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GenderCode")));
		Select s2=new Select(Gender);
		s2.selectByIndex(GenderCode);
		if(GenderCode==2)
		{
		WebElement Height = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PhysicalDetail_Height")));
		Height.sendKeys(PhysicalDetail_Height);
		}
		else {
		WebElement Height = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PhysicalDetail_Height")));
		Height.sendKeys(PhysicalDetail_Height);
		
		WebElement weight = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PhysicalDetail_Weight")));
		weight.sendKeys(PhysicalDetail_Weight);
		}
		
		//For Postal Address Details
		
		WebElement Door = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_DoorNo")));
		Door.sendKeys(DoorNo);
		
		WebElement Street = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_Street")));
		Street.sendKeys(ContactAddress_Street);
		
		WebElement Taluk = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_Taluk")));
		Taluk.sendKeys(ContactAddress_Taluk);
		
		WebElement City = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_City")));
		City.sendKeys(ContactAddress_City);
	
		WebElement State = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_UnionStateCode")));
	
		Select s3=new Select(State);
		Thread.sleep(1000);
		s3.selectByIndex(UnionStateCode);
		r.keyPress(KeyEvent.VK_TAB);
		r.keyRelease(KeyEvent.VK_TAB);
		
		WebElement District = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_DistrictCode")));
		Select s4=new Select(District);
		Thread.sleep(1000);
		s4.selectByIndex(DistrictCode);
		
		WebElement pincode = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_Pincode")));
		pincode.sendKeys(Address_Pincode);
		
		WebElement landmark = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_Landmark")));
		landmark.sendKeys(Address_Landmark);
		
		WebElement Native = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_NativeDistrict")));
		Native.sendKeys(NativeDistrict);
		
		WebElement permanent = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_ContactAddress_IsPermanentAddressSame' and @ value='False']")));
		act.moveToElement(permanent).click().perform();
		
		jss.executeScript("window.scrollBy(0,400)", "");
		
		//For Permanent Address Details 

		WebElement Doorp = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PermanentAddress_DoorNo")));
		Doorp.sendKeys(PermanentDoorNo);
		
		WebElement Streetp = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PermanentAddress_Street")));
		Streetp.sendKeys(PermanentAddress_Street);
		
		WebElement Talukp = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PermanentAddress_Taluk")));
		Talukp.sendKeys(PermanentAddress_Taluk);
		
		WebElement Cityp = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PermanentAddress_City")));
		Cityp.sendKeys(PermanentAddress_City);
		
		WebElement Statep = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PermanentAddress_UnionStateCode")));
		Select s5=new Select(Statep);
		Thread.sleep(500);
		s5.selectByIndex(PermanentUnionStateCode);
		r.keyPress(KeyEvent.VK_TAB);
		r.keyRelease(KeyEvent.VK_TAB);
		
		
		WebElement Districtp = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PermanentAddress_DistrictCode")));
		Select s6=new Select(Districtp);
		Thread.sleep(500);
		s6.selectByIndex(PermanentDistrictCode);
		
		WebElement pincodep = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PermanentAddress_Pincode")));
		pincodep.sendKeys(Permanent_Pincode);
		
		WebElement landmarkp = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PermanentAddress_Landmark")));
		landmarkp.sendKeys(Permanent_Landmark);
		jss.executeScript("window.scrollBy(0,400)", "");
		
		//For Reservation Details
		
		WebElement catagory = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_CategoryCode")));
		Select s7=new Select(catagory);
		s7.selectByIndex(CategoryCode);
		
		WebElement SubCaste = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_SubCaste")));
		SubCaste.sendKeys(Reservation_SubCaste);
		Thread.sleep(1000);

        Thread.sleep(500);
        WebElement datePicker1 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='Applicant_Reservation_CategoryCertificateIssuedDate']/following-sibling::input[1]")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", datePicker1);
        Dob(CategoryCertificateIssuedDate, driver, "(//input[@class='numInput cur-year'])[2]", "(//select[contains(@class,'flatpickr-monthDropdown-months')])[2]");
        // Select the year

//          
        WebElement ApplicantSportsAchivements = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantSportsAchivements_SportsCode")));
        Select s1=new Select(ApplicantSportsAchivements);
        s1.selectByIndex(Sportsachivement);
        jss.executeScript("window.scrollBy(0,400)", "");
        
        WebElement SportsAchivementCode = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantSportsAchivements_SportsAchivementCode")));
        Select s10 = new Select(SportsAchivementCode);
        s10.selectByIndex(game);
		
        Thread.sleep(500);
        WebElement gameachivementdeatils = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantSportsAchivements_AchivementDetails")));
        gameachivementdeatils.sendKeys(gamedetails);
        jss.executeScript("window.scrollBy(0,400)", "");
        
        WebElement ClaimingPDPReservation = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsClaimingPDPReservation' and @value='True']")));
        Thread.sleep(500);
        act.moveToElement(ClaimingPDPReservation).click().perform();
        
        WebElement KannadaMediumReservation = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsClamingKannadaMediumReservation' and @value='True']")));
        act.moveToElement(KannadaMediumReservation).click().perform();
        Thread.sleep(500);
        WebElement RuralReservation = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsClaimingRuralReservation' and @value='True']")));
        act.moveToElement(RuralReservation).click().perform();
        jss.executeScript("window.scrollBy(0,300)", "");
        
        WebElement KalyanaKarnataka = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsBelongToKalyanaKarnataka' and @value='False']")));
        act.moveToElement(KalyanaKarnataka).click().perform();
        Thread.sleep(500);
        WebElement KalyanaKarnatakaReservation = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsClamingKalyanaKarnatakaReservation' and @value='True']")));
        act.moveToElement(KalyanaKarnatakaReservation).click().perform();
        

        Thread.sleep(500);
        WebElement datePicker2 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='Applicant_Reservation_ApplyingforKalyanaKarnatakaCertificateDate']/following-sibling::input[1]")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", datePicker2);
        Dob(KKCertificateIssuedDate, driver, "(//input[@class='numInput cur-year'])[6]", "(//select[contains(@class,'flatpickr-monthDropdown-months')])[6]");
        
        
     
        Thread.sleep(500);
        WebElement GovernmentEmployee = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_AreYouAGovernmentEmployee' and @value='True']")));
        act.moveToElement(GovernmentEmployee).click().perform();
        
        WebElement datePicker3 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='Applicant_Reservation_GovermentServiceDetail_JoiningDate']/following-sibling::input[1]")));
     ((JavascriptExecutor) driver).executeScript("arguments[0].click();", datePicker3);
        Dob(joining_date_service, driver,"(//input[@class='numInput cur-year'])[5] ", "(//select[contains(@class,'flatpickr-monthDropdown-months')])[5]" );
    
     WebElement ser_year  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GovermentServiceDetail_YearsInService")));
 	Select s14=new Select(ser_year );
 	s14.selectByIndex(YearsInService_govt);
 	
 	WebElement ser_month  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GovermentServiceDetail_MonthsInService")));
 	Select s15=new Select(ser_month );
 	s15.selectByIndex(MonthsInService_govt);
 	
 	WebElement ser_day  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GovermentServiceDetail_DaysInService")));
 	Select s16=new Select(ser_day );
 	s16.selectByIndex(DaysInService_govt);	

 	WebElement GovtDept = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GovermentServiceDetail_Department")));
 	GovtDept.sendKeys(GovermentServiceDetail_Department);
 	
 	WebElement Designation = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GovermentServiceDetail_Designation")));
 	Designation.sendKeys(GovermentServiceDetail_Designation);
 	
 	WebElement DeptEnquiry = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_GovermentServiceDetail_HasDepartmentEnquiry' and @value='True']")));
 	act.moveToElement(DeptEnquiry).click().perform();
 
 	Thread.sleep(200);
 	
 	WebElement DeptEnquirydetails = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GovermentServiceDetail_DepartmentEnquiryDetail")));
 	DeptEnquirydetails.sendKeys(DepartmentEnquiryDetail);
	 jss.executeScript("window.scrollBy(0,300)", "");   
 	if(applyingpost==2)
 	{
 		WebElement KSp  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_applicantInService_KSPWingCode")));
 	 	Select s17=new Select(KSp );
 	 	s17.selectByIndex(KSP);	
 	 	
 	 	WebElement Designationins = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_applicantInService_Designation")));
 	 	Designationins.sendKeys(applicantInService_Designation);
 	 	
 	 	
 	}
 	
 	
 	
        }	
	}
				}
				    
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
    //  driver.quit(); // Close the WebDriver session
      
    }
    }
}
	private static void Dob(String Date, WebDriver driver, String YearXPath, String MonthXPath) throws InterruptedException {
	    String[] dateParts3 = Date.split("-");
	    String day3 = String.valueOf(Integer.parseInt(dateParts3[0])); // Remove leading zeros
	    String month3 = dateParts3[1];
	    String year3 = dateParts3[2];

	    WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

	    // Select the year
	    WebElement yearInput3 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(YearXPath)));
	    ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", yearInput3);
	    ((JavascriptExecutor) driver).executeScript("arguments[0].value='" + year3 + "';", yearInput3);
	    Thread.sleep(500);
	    yearInput3.sendKeys(Keys.ENTER);

	    // Select the month
	    WebElement monthDropdown4 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(MonthXPath)));
	    monthDropdown4.click();
	    Select monthSelect4 = new Select(monthDropdown4);
	    monthSelect4.selectByVisibleText(month3);
	    Thread.sleep(500);

	    // Select the day
	    String dateToSelect = month3 + " " + day3 + ", " + year3;
	    String dayXpath = "//span[@class='flatpickr-day' and @aria-label='" + dateToSelect + "']";
	    
	    WebElement dayElement3 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(dayXpath)));
	    ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", dayElement3);
	    ((JavascriptExecutor) driver).executeScript("arguments[0].click();", dayElement3);
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
     SimpleDateFormat sdf = new SimpleDateFormat("dd-MMMM-yyyy"); // Customize format
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

	  public static void selectDate3(WebDriver driver, String day3, String month3, String year3) {
	        // Method stub for selecting a date
	    }
	private void switchToNewWindow(WebDriver driver) {
	Set<String> windowHandles = driver.getWindowHandles();
	for (String windowHandle : windowHandles) {
	    driver.switchTo().window(windowHandle);
	}}
				

	}

