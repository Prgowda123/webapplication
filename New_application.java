package psi_sports;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Reporter;
import org.testng.annotations.Test;

public class New_application {
 
	@Test
	public void sample() throws InterruptedException, AWTException, ParseException
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

				
				
				
				
				for (int i = 2; i <= 34; i++) { // Start from row 1 to skip header
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
			int YearsInService =  (int) row.getCell(69).getNumericCellValue();
			int MonthsInService =  (int) row.getCell(70).getNumericCellValue();
			int DaysInService =  (int) row.getCell(71).getNumericCellValue();
			String criminal_case = getCellValue(row.getCell(73));
			String convicated_case = getCellValue(row.getCell(75));
			int year1 =  (int) row.getCell(77).getNumericCellValue();
			int year2 =  (int) row.getCell(78).getNumericCellValue();
			int year3 =  (int) row.getCell(79).getNumericCellValue();
			String DomiciledFromDate = getCellValue(row.getCell(80));
			String DomiciledToDate = getCellValue(row.getCell(81));
			
			int QualificationBoardCode =  (int) row.getCell(84).getNumericCellValue();
			String otherboard = getCellValue(row.getCell(85));
			int KannadaLanguagePaper =  (int) row.getCell(86).getNumericCellValue();
			int passingyear =  (int) row.getCell(87).getNumericCellValue();
			String m_mark = getCellValue(row.getCell(89));
			String ob_mark = getCellValue(row.getCell(90));
			String gradess = getCellValue(row.getCell(92));
			String perss = getCellValue(row.getCell(93));
			String RegistrationNo = getCellValue(row.getCell(94));
			
			int QualificationBoardCodepu =  (int) row.getCell(96).getNumericCellValue();
			String OtherBoardName = getCellValue(row.getCell(97));
			int pucPassingyear =  (int) row.getCell(98).getNumericCellValue();
			String  pu_max = getCellValue(row.getCell(100));
			String pu_ob = getCellValue(row.getCell(101));
			String  Grade_Grade = getCellValue(row.getCell(103));
			String ScorePercentage = getCellValue(row.getCell(104));
			String puRegistrationNo = getCellValue(row.getCell(105));
			
			String WhichDegree = getCellValue(row.getCell(107));
			String university = getCellValue(row.getCell(108));
			int degreePassingyear = (int)row.getCell(109).getNumericCellValue();
			String DegreeRegistrationNo = getCellValue(row.getCell(110));
			String  degree_max = getCellValue(row.getCell(112));
			String degree_ob = getCellValue(row.getCell(113));
			String  degree_Grade = getCellValue(row.getCell(115));
			String DScorePercentage = getCellValue(row.getCell(116));
			String Resultdate = getCellValue(row.getCell(117));
			int degreeEducationModeCode = (int)row.getCell(118).getNumericCellValue();
				
			int IdentityCardTypeCode =  (int) row.getCell(119).getNumericCellValue();
			String  UploadedIDNo = getCellValue(row.getCell(120));
			String IdentificationMark_01 = getCellValue(row.getCell(121));
			String  IdentificationMark_02 = getCellValue(row.getCell(122));
			String Applicant_Photo = getCellValue(row.getCell(123));
			String  Applicant_Signature = getCellValue(row.getCell(124));
			String Applicant_Thumb = getCellValue(row.getCell(125));
			String  Applicant_IdentityCard = getCellValue(row.getCell(126));
			String Meritotious_Sports_Certificate01  = getCellValue(row.getCell(127));
			String  Meritotious_Sports_Certificate02  = getCellValue(row.getCell(128));
			
			 
			
		WebElement ApplyingType = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplyingTypeCode")));
		Select s=new Select(ApplyingType);
		s.selectByIndex(applyingpost);
		
		WebElement prioriti = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantPostUnitsPriority_0__PostUnitCode")));
		Select p=new Select(prioriti);
		p.selectByIndex(Priority);

		WebElement prioriti1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantPostUnitsPriority_1__PostUnitCode")));
		Select p1=new Select(prioriti1);
		p1.selectByIndex(Priority1);
		
		WebElement prioriti2 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantPostUnitsPriority_2__PostUnitCode")));
		Select p2=new Select(prioriti2);
		p2.selectByIndex(Priority2);
		
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
		Thread.sleep(100);
		
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
		Thread.sleep(500);
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
		
		Thread.sleep(500);
		if(CategoryCode!=6)
		{
		WebElement SubCaste = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_SubCaste")));
		SubCaste.sendKeys(Reservation_SubCaste);
		Thread.sleep(500);
		
        WebElement datePicker1 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='Applicant_Reservation_CategoryCertificateIssuedDate']/following-sibling::input[1]")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", datePicker1);
        Dob(CategoryCertificateIssuedDate, driver, "(//input[@class='numInput cur-year'])[2]", "(//select[contains(@class,'flatpickr-monthDropdown-months')])[2]");
        // Select the year
		}
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
        Thread.sleep(500);
        act.moveToElement(RuralReservation).click().perform();
        jss.executeScript("window.scrollBy(0,300)", "");
        
        if(applyingpost==2)
        {
        WebElement KalyanaKarnataka = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsBelongToKalyanaKarnataka' and @value='False']")));
        Thread.sleep(500);
        act.moveToElement(KalyanaKarnataka).click().perform();
        Thread.sleep(500);
        WebElement KalyanaKarnatakaReservation = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsClamingKalyanaKarnatakaReservation' and @value='True']")));
        Thread.sleep(800);
        act.moveToElement(KalyanaKarnatakaReservation).click().perform();
        

        Thread.sleep(500);
        WebElement datePicker2 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='Applicant_Reservation_ApplyingforKalyanaKarnatakaCertificateDate']/following-sibling::input[1]")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", datePicker2);
        Dob(KKCertificateIssuedDate, driver, "(//input[@class='numInput cur-year'])[6]", "(//select[contains(@class,'flatpickr-monthDropdown-months')])[6]");
        
        }
        else
        {
        	  WebElement KalyanaKarnataka = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsBelongToKalyanaKarnataka' and @value='True']")));
              act.moveToElement(KalyanaKarnataka).click().perform();	
        }
     
        Thread.sleep(500);
        WebElement GovernmentEmployee = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_AreYouAGovernmentEmployee' and @value='True']")));
        act.moveToElement(GovernmentEmployee).click().perform();
        
        WebElement datePicker3 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='Applicant_Reservation_GovermentServiceDetail_JoiningDate']/following-sibling::input[1]")));
     ((JavascriptExecutor) driver).executeScript("arguments[0].click();", datePicker3);
        Dob(joining_date, driver,"(//input[@class='numInput cur-year'])[5] ", "(//select[contains(@class,'flatpickr-monthDropdown-months')])[5]" );
    
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
 	 	
 	   Thread.sleep(500);
       WebElement datePicker4 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='Applicant_applicantInService_JoiningDate']/following-sibling::input[1]")));
       ((JavascriptExecutor) driver).executeScript("arguments[0].click();", datePicker4);

       
       WebElement inser_year  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_applicantInService_YearsInService")));
    	Select s18=new Select(inser_year );
    	s18.selectByIndex(YearsInService);
    	
    	WebElement inser_month  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_applicantInService_MonthsInService")));
    	Select s19=new Select(inser_month );
    	s19.selectByIndex(MonthsInService);
    	
    	WebElement inser_day  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_applicantInService_DaysInService")));
    	Select s20=new Select(inser_day );
    	s20.selectByIndex(DaysInService);	
 	 	
 	 	
 	}
 	
 	Thread.sleep(500);
    WebElement CriminalActivity = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_CriminalActivity_IsInvolvedInCriminalActivity' and @value='True']")));
    act.moveToElement(CriminalActivity).click().perform();
    
    WebElement CriminalActivityCaseDetail = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_CriminalActivity_CaseDetail")));
   CriminalActivityCaseDetail.sendKeys(criminal_case);
    
   Thread.sleep(500);
   WebElement ConvictedInCriminalCase = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_CriminalActivity_IsConvictedInCriminalCase' and @value='True']")));
   act.moveToElement(ConvictedInCriminalCase).click().perform();
   
   WebElement ConvictionDetail = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_CriminalActivity_ConvictionDetail")));
   ConvictionDetail.sendKeys(convicated_case);
   
   Thread.sleep(500);
   WebElement DomiciledInKarnataka = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_ApplicantSportsAchivements_IsDomiciledInKarnataka' and @value='True']")));
   act.moveToElement(DomiciledInKarnataka).click().perform();
   jss.executeScript("window.scrollBy(0,300)", ""); 
   
   Thread.sleep(500);
   WebElement year01  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantSportsAchivements_RepresentingKarnatakaYears1")));
	Select s21=new Select(year01);
	s21.selectByIndex(year1);
	
	WebElement year02  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantSportsAchivements_RepresentingKarnatakaYears2")));
	Select s22=new Select(year02 );
	s22.selectByIndex(year2);
	
	WebElement year03  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplicantSportsAchivements_RepresentingKarnatakaYears3")));
	Select s20=new Select(year03 );
	s20.selectByIndex(year3);	
	 
	   Thread.sleep(500);
       WebElement datePicker5 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='Applicant_ApplicantSportsAchivements_DomiciledFromDate']/following-sibling::input[1]")));
       ((JavascriptExecutor) driver).executeScript("arguments[0].click();", datePicker5);
       Dob(DomiciledFromDate, driver, "(//input[@class='numInput cur-year'])[7]", "(//select[contains(@class,'flatpickr-monthDropdown-months')])[7]");
	
       WebElement datePicker6 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='Applicant_ApplicantSportsAchivements_DomiciledToDate']/following-sibling::input[1]")));
       ((JavascriptExecutor) driver).executeScript("arguments[0].click();", datePicker6);
       Dob(DomiciledToDate, driver, "(//input[@class='numInput cur-year'])[8]", "(//select[contains(@class,'flatpickr-monthDropdown-months')])[8]");
       
       
       Thread.sleep(500);
       WebElement Domiciled_Certificate = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_ApplicantSportsAchivements_IsObtainedDomiciledCert' and @value='True']")));
       act.moveToElement(Domiciled_Certificate).click().perform();
       
     //For Educational Qualification Details SSLC
   	
   	WebElement SSLC = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_IsSSLCHolder' and @ value='True']")));
   	Thread.sleep(600);
       act.moveToElement(SSLC).click().perform();
   	
   	jss.executeScript("window.scrollBy(0,300)", "");
   	Thread.sleep(100);
   	WebElement Board  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_QualificationBoardCode")));
   	Select s17=new Select(Board );
   	s17.selectByIndex(QualificationBoardCode);
   	
   	if(QualificationBoardCode==4)
   	{
   		WebElement OtherBoard = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_OtherBoardName")));
   		OtherBoard.sendKeys(otherboard);
   	}
   	
   	WebElement paper  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_KannadaLanguagePaper")));
   	Select s18=new Select(paper );
   	s18.selectByIndex(KannadaLanguagePaper);
   	
   	Thread.sleep(100);

   	WebElement Passingyear  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_YearOfPassing")));
   	Passingyear.click();
   	Thread.sleep(200);
   	Select s25=new Select(Passingyear );
   	Thread.sleep(200);
   	s25.selectByIndex(passingyear);
   	
   	if(applyingpost == 2)
   	{

   	
   	WebElement Markorgrade = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_SSLCQualification_MarkType' and @ value='M']")));
   	act.moveToElement(Markorgrade).click().perform();
   	
   	//convert from double to int
   	double doubleValue = Double.parseDouble(m_mark.trim()); // Handles "625.0"
   	int intValue = (int) doubleValue; // Converts to 625
   	// Convert the integer back to String for sendKeys
   	String intAsString = String.valueOf(intValue);

   	WebElement Maxmark = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_Score_Maximum")));
   	Maxmark.sendKeys(intAsString);
   	
   	double doubleValue1 = Double.parseDouble(ob_mark.trim()); // Handles "625.0"
   	int intValue1 = (int) doubleValue1; // Converts to 625
   	String intAsString1 = String.valueOf(intValue1);

   	WebElement Obtainmark = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_Score_Obtained")));
   	Obtainmark.sendKeys(intAsString1);
   	}
   	else {
   		WebElement Markorgrade = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_SSLCQualification_MarkType' and @ value='G']")));
   		act.moveToElement(Markorgrade).click().perform();
   		
   		
   		WebElement Gradesslc = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_Grade_Grade")));
   		Gradesslc.sendKeys(gradess);
   		
   		WebElement Persslc = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_ScoredPercentage")));
   		Persslc.sendKeys(perss);
   	}
   	WebElement Regno = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_RegistrationNo")));
   	Regno.sendKeys(RegistrationNo);
   	
	//For Educational Qualification Details PUC
   	Thread.sleep(200);
	WebElement Puc = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_IsPUCHolder' and @ value='True']")));
	Thread.sleep(200);
	act.moveToElement(Puc).click().perform();
	
	jss.executeScript("window.scrollBy(0,300)", "");
	
	WebElement PUBoard  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_QualificationBoardCode")));
	Select s26=new Select(PUBoard );
	s26.selectByIndex(QualificationBoardCodepu);
	
	if(QualificationBoardCodepu==4)
	{
		WebElement PUOtherBoard = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_OtherBoardName")));
		PUOtherBoard.sendKeys(OtherBoardName);
	}
	WebElement PUPassingyear  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_YearOfPassing")));
	Select s27=new Select(PUPassingyear);
	s27.selectByIndex(pucPassingyear);
	Thread.sleep(200);
	if(applyingpost==1) {
		Thread.sleep(200);
	WebElement PUMarkorgrade = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_PUCQualification_MarkType' and @ value='G']")));
	Thread.sleep(200);
	act.moveToElement(PUMarkorgrade).click().perform();
	
	WebElement Grade = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_Grade_Grade")));
	Grade.sendKeys(Grade_Grade);
	
	WebElement Per = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_ScorePercentage")));
	Per.sendKeys(ScorePercentage);
		}
	
	else
	{
		Thread.sleep(200);
		WebElement PUMarkorgrade = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_PUCQualification_MarkType' and @ value='M']")));
		Thread.sleep(200);
		act.moveToElement(PUMarkorgrade).click().perform();
		
		double doubleValue = Double.parseDouble(pu_max.trim()); // Handles "625.0"
		int intValue = (int) doubleValue; // Converts to 625
		// Convert the integer back to String for sendKeys
		String intAsString = String.valueOf(intValue);

		WebElement Maxmark = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_Score_Maximum")));
		Maxmark.sendKeys(intAsString);
		
		double doubleValue1 = Double.parseDouble(pu_ob.trim()); // Handles "625.0"
		int intValue1 = (int) doubleValue1; // Converts to 625
		String intAsString1 = String.valueOf(intValue1);

		WebElement Obtainmark = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_Score_Obtained")));
		Obtainmark.sendKeys(intAsString1);
	}
	WebElement PURegno = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_RegistrationNo")));
	PURegno.sendKeys(puRegistrationNo);
	
	//For Degree Details
	Thread.sleep(300);

	WebElement Degree = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_IsDegreeHolder' and @ value='True']")));
	act.moveToElement(Degree).click().perform();
	jss.executeScript("window.scrollBy(0,400)", "");
	
   	
	WebElement DegreeQualification = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_DegreeQualification_Degree")));
	DegreeQualification.sendKeys(WhichDegree);
	
	WebElement University = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_DegreeQualification_University")));
	University.sendKeys(university);
       
	WebElement DegreePassingyear  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_DegreeQualification_PassingMoY")));
	Select s28=new Select(DegreePassingyear);
	s28.selectByIndex(degreePassingyear);
	
	
	WebElement degreeRegistrationNo = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_DegreeQualification_RegistrationNo")));
	degreeRegistrationNo.sendKeys(DegreeRegistrationNo);
       
	Thread.sleep(200);
	if(applyingpost==2) {
	WebElement degreeMarkorgrade = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_DegreeQualification_MarkType' and @ value='G']")));
	act.moveToElement(degreeMarkorgrade).click().perform();
	
	WebElement Grade = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_DegreeQualification_Grade_Grade")));
	Grade.sendKeys(degree_Grade);
	
	WebElement Per = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_DegreeQualification_ScorePercentage")));
	Per.sendKeys(DScorePercentage);
		}
	
	else
	{
		Thread.sleep(200);
		WebElement DegreeMarkorgrade = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_DegreeQualification_MarkType' and @ value='M']")));
		Thread.sleep(200);
		act.moveToElement(DegreeMarkorgrade).click().perform();
		
		double doubleValue = Double.parseDouble(degree_max.trim()); // Handles "625.0"
		int intValue = (int) doubleValue; // Converts to 625
		// Convert the integer back to String for sendKeys
		String intAsString = String.valueOf(intValue);

		WebElement Maxmark = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_DegreeQualification_Score_Maximum")));
		Maxmark.sendKeys(intAsString);
		
		double doubleValue1 = Double.parseDouble(degree_ob.trim()); // Handles "625.0"
		int intValue1 = (int) doubleValue1; // Converts to 625
		String intAsString1 = String.valueOf(intValue1);

		WebElement Obtainmark = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_DegreeQualification_Score_Obtained")));
		Obtainmark.sendKeys(intAsString1);
	}
	 WebElement datePicker7 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='Applicant_EducationalQualification_DegreeQualification_ResultDate']/following-sibling::input[1]")));
     ((JavascriptExecutor) driver).executeScript("arguments[0].click();", datePicker7);
     Dob(Resultdate, driver, "(//input[@class='numInput cur-year'])[4]", "(//select[contains(@class,'flatpickr-monthDropdown-months')])[4]");
     
 	WebElement EducationModeCode  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_DegreeQualification_EducationModeCode")));
 	Select s29=new Select(EducationModeCode);
 	s29.selectByIndex(degreeEducationModeCode);
 	
	//For Documents Upload 
	WebElement IDcard = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_IdentityCardTypeCode")));
	Select s30=new Select(IDcard );
	s30.selectByIndex(IdentityCardTypeCode);
	
	
	if(IdentityCardTypeCode !=1) {
		WebElement IDcardno = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_UploadedIDNo")));
		IDcardno.sendKeys(UploadedIDNo);
	}
	
	WebElement Mark1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_IdentificationMark_01")));
	Mark1.sendKeys(IdentificationMark_01);
	
	WebElement Mark2 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_IdentificationMark_02")));
	Mark2.sendKeys(IdentificationMark_02);
	Thread.sleep(500);
	jss.executeScript("window.scrollBy(0,600)", "");
	 
       WebElement photo=wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Photo")));
       photo.sendKeys(Applicant_Photo);//
       
       WebElement sign=wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Signature")));
       sign.sendKeys(Applicant_Signature);
       
       WebElement thumb=wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Thumb")));
       thumb.sendKeys(Applicant_Thumb);
       
       WebElement ID=wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_IdentityCard")));
       ID.sendKeys(Applicant_IdentityCard);
 	
       WebElement merit1=wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_merit1")));
       merit1.sendKeys(Meritotious_Sports_Certificate01);
       
       WebElement merit2=wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_merit2")));
       merit2.sendKeys(Meritotious_Sports_Certificate02);
       jss.executeScript("window.scrollBy(0,600)", "");
       
       Thread.sleep(1000);
       WebElement preview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("preview-btn")));
       act.moveToElement(preview).click().perform();
       r.keyPress(KeyEvent.VK_ENTER);
	   r.keyRelease(KeyEvent.VK_ENTER);
       
       
	   //Preview contents
       Sheet sheet2 = workbook.getSheetAt(2);
	   Row row1 = sheet2.createRow(sheet2.getPhysicalNumberOfRows()); 
	   WebElement appdetails = driver.findElement(By.xpath("//h5[@id='exampleModalCenterTitle']"));
		String details1 = appdetails.getText();
		
		if(!isElementClickable(driver, appdetails)) {try {
		    boolean hasError = true;

		    // List of specific error IDs (extend this as needed)
		    String[] errorIds = {"Applicant_ApplicantPostUnitsPriority_0__PostUnitCode-error","Applicant_ApplicantPostUnitsPriority_1__PostUnitCode-error",
		    		"Applicant_ApplicantPostUnitsPriority_2__PostUnitCode-error","Applicant_ApplicantPostUnitsPriority_3__PostUnitCode-error","Applicant_ApplicantPostUnitsPriority_4__PostUnitCode-error","Applicant_ApplicantPostUnitsPriority_8__PostUnitCode-error",
		    		"Applicant_ApplicantPostUnitsPriority_5__PostUnitCode-error","Applicant_ApplicantPostUnitsPriority_6__PostUnitCode-error","Applicant_ApplicantPostUnitsPriority_7__PostUnitCode-error","Applicant_ApplicantPostUnitsPriority_9__PostUnitCode-error",
		    		"Applicant_ApplyingTypeCode-error","Applicant_PostUnitCode-error", "Applicant_FullName-error", "Applicant_FatherName-error", "Applicant_MotherName-error",
		        "Applicant_EmailId-error", "Applicant_MobileNo-error", "Applicant_AadharNo-error", "Applicant_DateOfBirth-error","Applicant_PhysicalDetail_Weight-error",
		        "Applicant_Reservation_GenderCode-error","Applicant_PhysicalDetail_Height-error","Applicant_ContactAddress_DoorNo-error","Applicant_ContactAddress_Street-error",
		        "Applicant_ContactAddress_Taluk-error","Applicant_ContactAddress_OtherDistrictName-error","Applicant_ContactAddress_Pincode",
		        "Applicant_ContactAddress_Landmark-error","Applicant_NativeDistrict-error","Applicant_PermanentAddress_DoorNo-error",
		        "Applicant_PermanentAddress_Street-error","Applicant_PermanentAddress_Taluk-error","Applicant_PermanentAddress_City-error",
		        "Applicant_PermanentAddress_OtherDistrictName-error","Applicant_PermanentAddress_Pincode","Applicant_PermanentAddress_Landmark-error",
		      "Applicant_Reservation_CategoryCode-error","Applicant_Reservation_SubCaste-error","Applicant_ApplicantSportsAchivements_SportsCode-error","Applicant_Reservation_GovermentServiceDetail_YearsInService-error","Applicant_Reservation_GovermentServiceDetail_Department-error",
		      "Applicant_ApplicantSportsAchivements_SportsAchivementCode-error","Applicant_ApplicantSportsAchivements_AchivementDetails-error","Applicant_applicantInService_KSPWingCode-error","Applicant_applicantInService_Designation-error",
		      "Applicant_applicantInService_YearsInService-error","Applicant_applicantInService_MonthsInService-error","Applicant_applicantInService_DaysInService-error","Applicant_ApplicantSportsAchivements_RepresentingKarnatakaYears1-error",
		      "Applicant_ApplicantSportsAchivements_RepresentingKarnatakaYears2-error","Applicant_ApplicantSportsAchivements_RepresentingKarnatakaYears3-error","Applicant_EducationalQualification_DegreeQualification_Degree-error",
		      "Applicant_EducationalQualification_DegreeQualification_University-error","Applicant_EducationalQualification_DegreeQualification_PassingMoY-error","Applicant_EducationalQualification_DegreeQualification_RegistrationNo-error","Applicant_EducationalQualification_DegreeQualification_Score_Maximum-error",
		      "Applicant_EducationalQualification_DegreeQualification_Score_Obtained-error","Applicant_EducationalQualification_DegreeQualification_ScorePercentage-error","Applicant_EducationalQualification_DegreeQualification_EducationModeCode-error","Applicant_EducationalQualification_DegreeQualification_Grade_Grade-error",
		      "Applicant_Reservation_GovermentServiceDetail_Designation-error","Applicant_CriminalActivity_DepartmentEnquiryDetail-error","Applicant_CriminalActivity_CaseDetail-error",
		      "Applicant_CriminalActivity_ConvictionDetail-error","Applicant_EducationalQualification_SSLCQualification_OtherBoardName-error","Applicant_EducationalQualification_SSLCQualification_Score_Maximum",
		      "Applicant_EducationalQualification_SSLCQualification_Score_Obtained-error","Applicant_EducationalQualification_SSLCQualification_Grade_Grade-error","Applicant_EducationalQualification_SSLCQualification_ScoredPercentage-error",
		      "Applicant_EducationalQualification_SSLCQualification_RegistrationNo-error","Applicant_EducationalQualification_PUCQualification_OtherBoardName-error","Applicant_EducationalQualification_PUCQualification_Score_Maximum-error",
		      "Applicant_EducationalQualification_PUCQualification_Score_Obtained-error","Applicant_EducationalQualification_PUCQualification_Grade_Grade-error","Applicant_EducationalQualification_PUCQualification_ScorePercentage-error",
		      "Applicant_EducationalQualification_PUCQualification_RegistrationNo-error","Applicant_IdentificationMark_01-error","Applicant_IdentificationMark_02-error",
		      "Applicant_CriminalActivity_ConvictionDetail-error"
		    };
		    
		    System.out.println(i + " : Iteration Failed");
            Reporter.log(i + " : Iteration Failed");

		 
		    WebDriverWait wait111 = new WebDriverWait(driver, Duration.ofMillis(10));
		    String ro = i+" Failed";
		    row1.createCell(1).setCellValue(ro);
		 
		    for (String errorId : errorIds) {
		        try {
		            WebElement errorMessage = wait111.until(ExpectedConditions.visibilityOfElementLocated(By.id(errorId)));
		            if (errorMessage != null && !errorMessage.getText().isEmpty()) {
		                hasError = false;
		                String er = errorMessage.getText();
		               
		                System.out.println("Error in field: " + errorId + " - " + er);
		           
		                Reporter.log(i + " iteration" + " Error in field of " + errorId + " - " + er);
		            }
		            if (!isElementClickable(driver, preview)) {
		                r.keyPress(KeyEvent.VK_ENTER);
		                r.keyRelease(KeyEvent.VK_ENTER);
		            }
		        } catch (TimeoutException e) {
		     
		//  System.out.println("No error for field: " + errorId);
		        }
		    }
		} catch (Exception q) {
		  //  q.printStackTrace();  
		}
		}
       
		//preview
		try
		{
			WebElement candidateTypePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='candidateTypePreview']")));
	       wait.until(ExpectedConditions.visibilityOf(candidateTypePreview));
			wait.until(ExpectedConditions.elementToBeClickable(candidateTypePreview));
	       String candidatePreview = candidateTypePreview.getText();
	       
	       WebElement priority1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@id='priorityPreview']//div[2]")));
	       wait.until(ExpectedConditions.visibilityOf(priority1));
			wait.until(ExpectedConditions.elementToBeClickable(priority1));
	       String Priority1Pr = priority1.getText();
	       
	       WebElement priority2 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@id='priorityPreview']//div[4]")));
	       wait.until(ExpectedConditions.visibilityOf(priority2));
			wait.until(ExpectedConditions.elementToBeClickable(priority2));
	       String Priority2Pr = priority2.getText();

	       WebElement priority3 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@id='priorityPreview']//div[6]")));
	       wait.until(ExpectedConditions.visibilityOf(priority3));
			wait.until(ExpectedConditions.elementToBeClickable(priority3));
	       String Priority3Pr = priority3.getText();
	     
	       WebElement priority4 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@id='priorityPreview']//div[8]")));
	       wait.until(ExpectedConditions.visibilityOf(priority4));
			wait.until(ExpectedConditions.elementToBeClickable(priority4));
	       String Priority4Pr = priority4.getText();
	
	       WebElement priority5 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@id='priorityPreview']//div[10]")));
	       wait.until(ExpectedConditions.visibilityOf(priority5));
			wait.until(ExpectedConditions.elementToBeClickable(priority5));
	       String Priority5Pr = priority5.getText();

	       WebElement priority6 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@id='priorityPreview']//div[12]")));
	       wait.until(ExpectedConditions.visibilityOf(priority6));
			wait.until(ExpectedConditions.elementToBeClickable(priority6));
	       String Priority6Pr = priority6.getText();
	
	       WebElement priority7 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@id='priorityPreview']//div[14]")));
	       wait.until(ExpectedConditions.visibilityOf(priority7));
			wait.until(ExpectedConditions.elementToBeClickable(priority7));
	       String Priority7Pr = priority6.getText();

	       WebElement priority8 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@id='priorityPreview']//div[16]")));
	       wait.until(ExpectedConditions.visibilityOf(priority8));
			wait.until(ExpectedConditions.elementToBeClickable(priority8));
	       String Priority8Pr = priority8.getText();
	  
	       WebElement priority9 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@id='priorityPreview']//div[18]")));
	       wait.until(ExpectedConditions.visibilityOf(priority9));
			wait.until(ExpectedConditions.elementToBeClickable(priority9));
	       String Priority9Pr = priority9.getText();
	     
	       WebElement priority10 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@id='priorityPreview']//div[20]")));
	       wait.until(ExpectedConditions.visibilityOf(priority10));
			wait.until(ExpectedConditions.elementToBeClickable(priority10));
	       String Priority10Pr = priority10.getText();
	    
	       row1.createCell(1).setCellValue(candidatePreview);
	       row1.createCell(2).setCellValue(Priority1Pr);
	       row1.createCell(3).setCellValue(Priority2Pr);
	       row1.createCell(4).setCellValue(Priority3Pr);
	       row1.createCell(5).setCellValue(Priority4Pr);
	       row1.createCell(6).setCellValue(Priority5Pr);
	       row1.createCell(7).setCellValue(Priority6Pr);
	       row1.createCell(8).setCellValue(Priority7Pr);
	       row1.createCell(9).setCellValue(Priority8Pr);
           row1.createCell(10).setCellValue(Priority9Pr);
           row1.createCell(11).setCellValue(Priority10Pr);
           
           WebElement candidatenameTypePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("candidatenameTypePreview")));
	       String candidatenamePreview = candidatenameTypePreview.getText();
	   
	       
	       WebElement fatherNamePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("fatherNamePreview")));
	       String fatherPreview = fatherNamePreview.getText();
	     
	       WebElement MotherNamePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("MotherNamePreview")));
	       String MotherPreview = MotherNamePreview.getText();
	       
	       WebElement emailPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("emailPreview")));
	       String emailidPreview = emailPreview.getText();
	   
	       
	       WebElement MobileNoPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("MobileNoPreview")));
	       String MobilePreview = MobileNoPreview.getText();
	       
	       WebElement aadharPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("aadharPreview")));
	       String aadharnoPreview = aadharPreview.getText();
	       jss.executeScript("arguments[0].scrollIntoView(true);", aadharPreview);
	       
	       WebElement DateofBirthPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DateofBirthPreview")));
	       String DateofBirtPreview = DateofBirthPreview.getText();
	       
	       WebElement DateofBirthasonPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DateofBirthasonPreview")));
	       String DateofBirthasonnPreview = DateofBirthasonPreview.getText();
	       
	       WebElement genderPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("genderPreview")));
	       String genderrPreview = genderPreview.getText();
	       
	       WebElement HeightPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("HeightPreview")));
	       String HeighttPreview = HeightPreview.getText();
	       
	       WebElement WeightPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("WeightPreview")));
	       String WeighttPreview = HeightPreview.getText();
	       jss.executeScript("arguments[0].scrollIntoView(true);", WeightPreview);	
	 
	       WebElement DoorPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DoorPreview")));
	       String DoornoPreview = DoorPreview.getText();
	       
	       WebElement StreetPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("StreetPreview")));
	       String StreettPreview = StreetPreview.getText();
	       
	       WebElement talukPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("talukPreview")));
	       String talukkPreview = talukPreview.getText();
	       
	       WebElement cityPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("cityPreview")));
	       String citiPreview = cityPreview.getText();
	       
	       WebElement statePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("statePreview")));
	       String stateePreview = statePreview.getText();
	       
	       WebElement districtPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("districtPreview")));
	       String districttPreview = districtPreview.getText();
	   
	       
	       WebElement pincodePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("pincodePreview")));
	       String pincodeePreview = pincodePreview.getText();
	       
	       WebElement landmarkPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("landmarkPreview")));
	       String landmarkkPreview = landmarkPreview.getText();
	       jss.executeScript("arguments[0].scrollIntoView(true);", landmarkPreview);
	       
	       WebElement nativeDistrictPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("nativeDistrictPreview")));
	       String nativeDistrictP = nativeDistrictPreview.getText();
	       
	       WebElement postaladdressPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("postaladdressPreview")));
	       String postaladdresPreview = postaladdressPreview.getText();
	       
	       WebElement PerDoorNoPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PerDoorNoPreview")));
	       String PermDoorNoPreview = PerDoorNoPreview.getText();
	       
	       WebElement PerStreetPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PerStreetPreview")));
	       String PermStreetPreview = PerStreetPreview.getText();
	       
	       WebElement PerTalukPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PerTalukPreview")));
	       String PermTalukPreview = PerTalukPreview.getText();
	       jss.executeScript("arguments[0].scrollIntoView(true);", PerTalukPreview);	
	 

	       WebElement PerCityPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PerCityPreview")));
	       String PermCityPreview = PerCityPreview.getText();
	       
	       WebElement perstatePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("perstatePreview")));
	       String permstatePreview = perstatePreview.getText();
	       
	       WebElement PerDistrictPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PerDistrictPreview")));
	       String PermDistrictPreview = PerDistrictPreview.getText();
	       
	       WebElement PerPincodePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PerPincodePreview")));
	       String PermPincodePreview = PerPincodePreview.getText();
	       
	       WebElement NearbyLandmarkPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("NearbyLandmarkPreview")));
	       String NearbyLandmarkkPreview = NearbyLandmarkPreview.getText();
	       
	       row1.createCell(12).setCellValue(candidatenamePreview);
	       row1.createCell(13).setCellValue(fatherPreview);
	       row1.createCell(14).setCellValue(MotherPreview);
	       row1.createCell(15).setCellValue(emailidPreview);
	       row1.createCell(16).setCellValue(MobilePreview);
	       row1.createCell(17).setCellValue(aadharnoPreview);
	       row1.createCell(18).setCellValue(DateofBirtPreview);
           row1.createCell(19).setCellValue(DateofBirthasonnPreview);
           row1.createCell(20).setCellValue(genderrPreview);
           row1.createCell(21).setCellValue(HeighttPreview);
           row1.createCell(22).setCellValue(WeighttPreview);
           row1.createCell(23).setCellValue(DoornoPreview);
           row1.createCell(24).setCellValue(StreettPreview);
           row1.createCell(25).setCellValue(talukkPreview);
           row1.createCell(26).setCellValue(citiPreview);
           row1.createCell(27).setCellValue(stateePreview);
           row1.createCell(28).setCellValue(districttPreview);
           row1.createCell(30).setCellValue(pincodeePreview);
           row1.createCell(31).setCellValue(landmarkkPreview);
           row1.createCell(32).setCellValue(nativeDistrictP);
           row1.createCell(33).setCellValue(postaladdresPreview);
           row1.createCell(34).setCellValue(PermDoorNoPreview);
           row1.createCell(35).setCellValue(PermStreetPreview);
           row1.createCell(36).setCellValue(PermTalukPreview);
           row1.createCell(37).setCellValue(PermCityPreview);
           row1.createCell(38).setCellValue(permstatePreview);
           row1.createCell(39).setCellValue(PermDistrictPreview);
           row1.createCell(41).setCellValue(PermPincodePreview);
           row1.createCell(42).setCellValue(NearbyLandmarkkPreview);
           
           //reservation
	       WebElement CategoryPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("CategoryPreview")));
	       String CategoriPreview = CategoryPreview.getText();
	   
	       
	       WebElement SubcastePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SubcastePreview")));
	       String SubcasteePreview = SubcastePreview.getText();
	       
	       WebElement DateofSubcastePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DateofSubcastePreview")));
	       String DateofSubcasteePreview = DateofSubcastePreview.getText();
	       jss.executeScript("arguments[0].scrollIntoView(true);", DateofSubcastePreview);
	       
	       row1.createCell(43).setCellValue(CategoriPreview);
           row1.createCell(44).setCellValue(SubcasteePreview);
           row1.createCell(45).setCellValue(DateofSubcasteePreview);
           
           WebElement WhichSportsPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("WhichSportsPreview")));
	       String whichSportsPreview = WhichSportsPreview.getText();
	   
	       
	       WebElement AchievementsPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("AchievementsPreview")));
	       String achievementsPreview = AchievementsPreview.getText();
           
	       WebElement sportsGamePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("sportsGamePreview")));
	       String SportsGamePreview = sportsGamePreview.getText();
	       
	       WebElement KannadaMediumReservationPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("KannadaMediumReservationPreview")));
	       String KannadaMediummReservationPreview = KannadaMediumReservationPreview.getText();
	       
	       WebElement PDPReservationPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PDPReservationPreview")));
	       String PDPReservationnPreview = PDPReservationPreview.getText();

	       WebElement ClaimingRuralMediumPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ClaimingRuralMediumPreview")));
	       String ClaimingRuralMediummPreview = ClaimingRuralMediumPreview.getText();
	       jss.executeScript("arguments[0].scrollIntoView(true);", ClaimingRuralMediumPreview);  
	       
	       row1.createCell(46).setCellValue(whichSportsPreview);
           row1.createCell(47).setCellValue(achievementsPreview);
           row1.createCell(48).setCellValue(SportsGamePreview);
	       row1.createCell(49).setCellValue(KannadaMediummReservationPreview);
           row1.createCell(50).setCellValue(PDPReservationnPreview);
           
           row1.createCell(51).setCellValue(ClaimingRuralMediummPreview); 
           
	       if(applyingpost == 2) {
	       WebElement KalyanaKarnatakaPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("KalyanaKarnatakaPreview")));
	       String KalyanaKarnatakaaPreview = KalyanaKarnatakaPreview.getText();
	   
	       WebElement KalyanaKarnatakaDistrictPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("KalyanaKarnatakaDistrictPreview")));
	       String KalyanaKarnatakaaDistrictPreview = KalyanaKarnatakaDistrictPreview.getText();
	       

	       WebElement modalkkcertificatedatepreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("modalkkcertificatedatepreview")));
	       String Modalkkcertificatedatepreview = modalkkcertificatedatepreview.getText();
	       
	       row1.createCell(52).setCellValue(KalyanaKarnatakaaPreview);
	       row1.createCell(53).setCellValue(KalyanaKarnatakaaDistrictPreview);
	       row1.createCell(54).setCellValue(Modalkkcertificatedatepreview);
	       }
	       else {
	    	   WebElement KalyanaKarnatakaPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("KalyanaKarnatakaPreview")));
		       String KalyanaKarnatakaaPreview = KalyanaKarnatakaPreview.getText();
		       row1.createCell(52).setCellValue(KalyanaKarnatakaaPreview);
	       }
	       
	       WebElement GovernmentEmployeePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("GovernmentEmployeePreview")));
	       String GovernmentEmployePreview = GovernmentEmployeePreview.getText();
	   
	       WebElement GovtDateofJoiningPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("GovtDateofJoiningPreview")));
	       String GovtDateofJoininggPreview = GovtDateofJoiningPreview.getText();
	       
	       WebElement GovtYearsPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("GovtYearsPreview")));
	       String GovtYearsPrevieww = GovtYearsPreview.getText();
	       
	       WebElement GovtMonthPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("GovtMonthPreview")));
	       String GovtMonthhPreview = GovtMonthPreview.getText();
	       
	       WebElement GovtDaysPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("GovtDaysPreview")));
	       String GovtDayPreview = GovtDaysPreview.getText();
	       jss.executeScript("arguments[0].scrollIntoView(true);", GovtDaysPreview);  
	       
	       WebElement GovtDepartmentPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("GovtDepartmentPreview")));
	       String GovtnDepartmentPreview = GovtDepartmentPreview.getText();
	       
	       WebElement GovtDesignationPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("GovtDesignationPreview")));
	       String GovtnDesignationPreview = GovtDesignationPreview.getText();
	       
	       WebElement DepartmentalEnquirPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DepartmentalEnquirPreview")));
	       String DepartmentalEnquiryPreview = DepartmentalEnquirPreview.getText();
	       
	       WebElement DeptenqdetailsPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DeptenqdetailsPreview")));
	       String DeptenqrydetailsPreview = DeptenqdetailsPreview.getText();
	       jss.executeScript("arguments[0].scrollIntoView(true);", DeptenqdetailsPreview);
	       
	       row1.createCell(55).setCellValue(GovernmentEmployePreview);
	       row1.createCell(56).setCellValue(GovtDateofJoininggPreview);
	       row1.createCell(58).setCellValue(GovtYearsPrevieww);
	       row1.createCell(59).setCellValue(GovtMonthhPreview);
	       row1.createCell(60).setCellValue(GovtDayPreview);
	       row1.createCell(61).setCellValue(GovtnDepartmentPreview);
	       row1.createCell(62).setCellValue(GovtnDesignationPreview);
	       row1.createCell(63).setCellValue(DepartmentalEnquiryPreview);
	       row1.createCell(64).setCellValue(DeptenqrydetailsPreview);
	       
	       
	       WebElement InservicePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("InservicePreview")));
	       String inservicePreview = InservicePreview.getText();
           
	       WebElement KspwingPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("KspwingPreview")));
	       String kspwingPreview = KspwingPreview.getText();
	   
	       WebElement inservicedesignnationPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("inservicedesignnationPreview")));
	       String InservicedesignnationPreview = inservicedesignnationPreview.getText();
	   
	       WebElement inservicedojPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("inservicedojPreview")));
	       String InservicedojPreview = inservicedojPreview.getText();
	       
	       WebElement inserviceYearsPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("inserviceYearsPreview")));
	       String InserviceYearsPreview = inserviceYearsPreview.getText();
	       
	       WebElement inserviceMonthPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("inserviceMonthPreview")));
	       String InserviceMonthPreview = inserviceMonthPreview.getText();
	       
	       WebElement inserviceDaysPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("inserviceDaysPreview")));
	       String InserviceDaysPreview = inserviceDaysPreview.getText();
	       jss.executeScript("arguments[0].scrollIntoView(true);", inserviceDaysPreview);  
           
	       row1.createCell(65).setCellValue(inservicePreview);
	       row1.createCell(66).setCellValue(kspwingPreview);
	       row1.createCell(67).setCellValue(InservicedesignnationPreview);
	       row1.createCell(68).setCellValue(InservicedojPreview);
	       row1.createCell(70).setCellValue(InserviceYearsPreview);
	       row1.createCell(71).setCellValue(InserviceMonthPreview);
	       row1.createCell(72).setCellValue(InserviceDaysPreview);
           
	       WebElement CriminalCasesPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("CriminalCasesPreview")));
	       String criminalCasesPreview = CriminalCasesPreview.getText();
           
	       WebElement CriminalCasesdetailsPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("CriminalCasesdetailsPreview")));
	       String criminalCasesdetailsPreview = CriminalCasesdetailsPreview.getText();
	   
	       WebElement ConvictedinaCriminalPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ConvictedinaCriminalPreview")));
	       String convictedinaCriminalPreview = ConvictedinaCriminalPreview.getText();
	   
	       WebElement ConvictedCriminalDetailsPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ConvictedCriminalDetailsPreview")));
	       String convictedCriminalDetailsPreview = ConvictedCriminalDetailsPreview.getText();
	       
	       WebElement DominicalKarnatakaapplicable = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DominicalKarnatakaapplicable")));
	       String Dominicalkarnatakaapplicable = DominicalKarnatakaapplicable.getText();
	       
	       WebElement Years01Preview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Years01Preview")));
	       String years01Preview = Years01Preview.getText();
	       
	       WebElement Years02Preview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Years02Preview")));
	       String years02Preview = Years02Preview.getText();
	       
	       WebElement Years03Preview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Years03Preview")));
	       String years03Preview = Years03Preview.getText();
	       
	       jss.executeScript("arguments[0].scrollIntoView(true);", Years03Preview);   
           
	       WebElement DominicalfromDatePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DominicalfromDatePreview")));
	       String dominicalfromDatePreview = DominicalfromDatePreview.getText();
           
	       WebElement DominicaltoDatePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DominicaltoDatePreview")));
	       String DominicalToDatePreview = DominicaltoDatePreview.getText();
	       
	       WebElement DominicalCalucation = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DominicalCalucation")));
	       String Dominicalcalucation = DominicalCalucation.getText();
	       jss.executeScript("arguments[0].scrollIntoView(true);", DominicalCalucation);   
	       WebElement DominicalCertificate = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DominicalCertificate")));
	       String Dominicalcertificate = DominicalCertificate.getText();
           
	       row1.createCell(73).setCellValue(criminalCasesPreview);
	       row1.createCell(74).setCellValue(criminalCasesdetailsPreview);
	       row1.createCell(75).setCellValue(convictedinaCriminalPreview);
	       row1.createCell(76).setCellValue(convictedCriminalDetailsPreview);
	       row1.createCell(77).setCellValue(Dominicalkarnatakaapplicable);
	       row1.createCell(78).setCellValue(years01Preview);
	       row1.createCell(79).setCellValue(years02Preview);
	       row1.createCell(80).setCellValue(years03Preview);
	       row1.createCell(81).setCellValue(dominicalfromDatePreview);
	       row1.createCell(82).setCellValue(DominicalToDatePreview);
	       row1.createCell(83).setCellValue(Dominicalcalucation);
	       row1.createCell(84).setCellValue(Dominicalcertificate);
           
	       
	 	     WebElement PassedSSLCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PassedSSLCPreview")));
		     String PasseSSLCPreview = PassedSSLCPreview.getText(); 
		         
		     WebElement BoardofSslcPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("BoardofSslcPreview")));
		     String BoardSslcPreview = BoardofSslcPreview.getText(); 
		     
		     row1.createCell(85).setCellValue(PasseSSLCPreview);
		     row1.createCell(86).setCellValue(BoardSslcPreview);
		       
		       
		      if(QualificationBoardCode==3)
		      {
		     WebElement OtherSslcBoarPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("OtherSslcBoarPreview")));
		     String OtherSslcBoardPreview = OtherSslcBoarPreview.getText(); 
		     row1.createCell(87).setCellValue(OtherSslcBoardPreview);
		       
		      }  
		     WebElement KannadaLanguagePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("KannadaLanguagePreview")));
		     String KannadaLanguagPreview = KannadaLanguagePreview.getText(); 
		        
	 	      
		     WebElement YearofPassingSSLCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("YearofPassingSSLCPreview")));
		     String YearfPassingSSLCPreview = YearofPassingSSLCPreview.getText(); 
		         
		     WebElement SSLCMarksorGradesPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SSLCMarksorGradesPreview")));
		     String SSLCMarksorGradePreview = SSLCMarksorGradesPreview.getText(); 
		     if(applyingpost ==2)  { 
		     WebElement SSLCMaxMarksPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SSLCMaxMarksPreview")));
		     String SSLCMaxMarkPreview = SSLCMaxMarksPreview.getText(); 
		      jss.executeScript("arguments[0].scrollIntoView(true);", SSLCMaxMarksPreview);   
	   
	 	      
		     WebElement SSLCMarksObtainedPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SSLCMarksObtainedPreview")));
		     String SSLCMarkObtainedPreview = SSLCMarksObtainedPreview.getText(); 
		     row1.createCell(91).setCellValue(SSLCMaxMarkPreview);
		     row1.createCell(92).setCellValue(SSLCMarkObtainedPreview);
		       
		     WebElement SSLCPercentageObtainedPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SSLCPercentageObtainedPreview")));
		     String SSLCPercentageObtainePreview = SSLCPercentageObtainedPreview.getText(); 
		     row1.createCell(93).setCellValue(SSLCPercentageObtainePreview);
		     }  
		     
		     else {
		    	   WebElement SSLCGradesObtainedPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SSLCGradesObtainedPreview")));
		  	     String SSLCGradesObtainePreview = SSLCGradesObtainedPreview.getText(); 
		  	     
		  	   WebElement SSLCPercentageObtainedPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SSLCPercentageObtainedPreview")));
			     String SSLCPercentageObtainePreview = SSLCPercentageObtainedPreview.getText(); 
			     row1.createCell(94).setCellValue(SSLCGradesObtainePreview);
			     row1.createCell(95).setCellValue(SSLCPercentageObtainePreview);
			     
		     }
		     WebElement SSLCRegistrationNoPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SSLCRegistrationNoPreview")));
		     String SSLCRegistrationsNoPreview = SSLCRegistrationNoPreview.getText(); 
		     
		     row1.createCell(88).setCellValue(KannadaLanguagPreview);
		     row1.createCell(89).setCellValue(YearfPassingSSLCPreview);
		     row1.createCell(90).setCellValue(SSLCMarksorGradePreview);
		     row1.createCell(96).setCellValue(SSLCRegistrationsNoPreview);
		        
		     //puc
	    WebElement PassedPUCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PassedPUCPreview")));
	    String PassePUCPreview = PassedPUCPreview.getText(); 
	    jss.executeScript("arguments[0].scrollIntoView(true);", PassedPUCPreview);  
	    
  
	    WebElement PassedPucBoardPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PassedPucBoardPreview")));
	    String PassePucBoardPreview = PassedPucBoardPreview.getText(); 
	    
	    row1.createCell(97).setCellValue(PassePUCPreview);
	    row1.createCell(98).setCellValue(PassePucBoardPreview);
	                 
	    if(QualificationBoardCodepu==3) {
	    WebElement PucOtherBoardPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PucOtherBoardPreview")));
	    String PuOtherBoardPreview = PucOtherBoardPreview.getText(); 
	    row1.createCell(99).setCellValue(PuOtherBoardPreview);
        
	    }
	    WebElement YearofPassingPUCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("YearofPassingPUCPreview")));
	    String YearPassingPUCPreview = YearofPassingPUCPreview.getText(); 
	    
	    WebElement MarksorGradePUCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("MarksorGradePUCPreview")));
	    String MarkorGradePUCPreview = MarksorGradePUCPreview.getText(); 
	    jss.executeScript("arguments[0].scrollIntoView(true);", MarksorGradePUCPreview); 
	    
	    if(applyingpost==1) {
	    WebElement GradesObtainedPUCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("GradesObtainedPUCPreview")));
	    String GradeObtainedPUCPreview = GradesObtainedPUCPreview.getText(); 
	    
	    WebElement PercentagePUCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PercentagePUCPreview")));
	    String PercentagePUPreview = PercentagePUCPreview.getText(); 
	    row1.createCell(105).setCellValue(GradeObtainedPUCPreview);
	    row1.createCell(106).setCellValue(PercentagePUPreview);
	    }
	    else {
		WebElement MaxMarksPUCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("MaxMarksPUCPreview")));
		String MaxMarksPUCPrevie = MaxMarksPUCPreview.getText(); 
		
		WebElement MarksobtainedPUCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("MarksobtainedPUCPreview")));
		String MarksobtainedPUCPrevie = MarksobtainedPUCPreview.getText(); 
	   
	     WebElement PercentagePUCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PercentagePUCPreview")));
	 	 String PercentagePUPreview = PercentagePUCPreview.getText(); 
	 	 row1.createCell(102).setCellValue(MaxMarksPUCPrevie);
	 	 row1.createCell(103).setCellValue(MarksobtainedPUCPrevie);
	 	 row1.createCell(104).setCellValue(PercentagePUPreview);
	    }
			    WebElement PUCRegistrationsnoPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PUCRegistrationsnoPreview")));
	    String PUCRegistrationnoPreview = PUCRegistrationsnoPreview.getText(); 
	    jss.executeScript("arguments[0].scrollIntoView(true);", PUCRegistrationsnoPreview); 
	    
	  row1.createCell(100).setCellValue(YearPassingPUCPreview);
	  row1.createCell(101).setCellValue(MarkorGradePUCPreview);
	 
	  row1.createCell(107).setCellValue(PUCRegistrationnoPreview);
	 
	 //DEgree
	 WebElement DegHolderPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DegHolderPreview")));
	 String DegreeHolderPreview = DegHolderPreview.getText(); 
	 row1.createCell(108).setCellValue(DegreeHolderPreview);
	    
	WebElement WhichDegreePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("WhichDegreePreview")));
	String whichDegreePreview = WhichDegreePreview.getText();
    
	WebElement WhichUnivercityPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("WhichUnivercityPreview")));
	String whichUnivercityPreview = WhichUnivercityPreview.getText();
	
	WebElement YearofPassingDegreePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("YearofPassingDegreePreview")));
	String YearofPassingdegreePreview = YearofPassingDegreePreview.getText();
    jss.executeScript("arguments[0].scrollIntoView(true);", YearofPassingDegreePreview);  
	  
	WebElement DegreeRegistrationNumberPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DegreeRegistrationNumberPreview")));
	String DegreeregistrationNumberPreview = DegreeRegistrationNumberPreview.getText();
	    
	WebElement MarksOrGradePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("MarksOrGradePreview")));
	String MarksorGradePreview = MarksOrGradePreview.getText();
    jss.executeScript("arguments[0].scrollIntoView(true);", MarksOrGradePreview);   
    
    
	  row1.createCell(109).setCellValue(whichDegreePreview);
      row1.createCell(110).setCellValue(whichUnivercityPreview);
      row1.createCell(111).setCellValue(YearofPassingdegreePreview);
      row1.createCell(112).setCellValue(DegreeregistrationNumberPreview);
      row1.createCell(113).setCellValue(MarksorGradePreview);
     
      if(applyingpost==2) {
  	    WebElement GradesObtainedPUCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("GradesObtainedDegPreview")));
  	    String GradeObtainedPUCPreview = GradesObtainedPUCPreview.getText(); 
  	    
  	    WebElement PercentagePUCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PercentageOrCGPAInDegreePreview")));
  	    String PercentagePUPreview = PercentagePUCPreview.getText(); 
  	    row1.createCell(117).setCellValue(GradeObtainedPUCPreview);
  	    row1.createCell(118).setCellValue(PercentagePUPreview);
  	    }
  	    else {
  		WebElement MaxMarksDegreePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("MaxMarksDegreePreview")));
  		String MaxMarksdegreePreview = MaxMarksDegreePreview.getText(); 
  		
  		WebElement MarksobtainedDegPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("MarksobtainedDegPreview")));
  		String MarksobtainedegPreview = MarksobtainedDegPreview.getText(); 
  	   
  	     WebElement PercentageOrCGPAInDegreePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PercentageOrCGPAInDegreePreview")));
  	 	 String PercentageorCGPAInDegreePreview = PercentageOrCGPAInDegreePreview.getText(); 
  	 	 row1.createCell(114).setCellValue(MaxMarksdegreePreview);
  	 	 row1.createCell(115).setCellValue(MarksobtainedegPreview);
  	 	 row1.createCell(116).setCellValue(PercentageorCGPAInDegreePreview);
  	    }     
		    
		     
		WebElement DateOfAnnouncementPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DateOfAnnouncementPreview")));
  		String DateOfannouncementPreview = DateOfAnnouncementPreview.getText(); 
  	   jss.executeScript("arguments[0].scrollIntoView(true);", DateOfAnnouncementPreview); 
  	  
  		WebElement degeduModePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("degeduModePreview")));
  		String degedumodePreview = degeduModePreview.getText();      
		     
         row1.createCell(119).setCellValue(DateOfannouncementPreview);
         row1.createCell(120).setCellValue(degedumodePreview);
         
         //document
 	    WebElement IDCardSelectedPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("IDCardSelectedPreview")));
 	    String IDCardSelectePreview = IDCardSelectedPreview.getText(); 
 	      
 	    WebElement SelectedIDCardNoPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SelectedIDCardNoPreview")));
 	    String SelectedIDCardNPreview = SelectedIDCardNoPreview.getText(); 
 	    jss.executeScript("arguments[0].scrollIntoView(true);", SelectedIDCardNoPreview); 
   
 	    WebElement Identitymark01Preview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Identitymark01Preview")));
 	    String Identitimark01Preview = Identitymark01Preview.getText(); 
 	      
 	    WebElement Identitymark02Preview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Identitymark02Preview")));
 	    String Identitimark02Preview = Identitymark02Preview.getText(); 
 	    jss.executeScript("arguments[0].scrollIntoView(true);", Identitymark02Preview);    
 	    row1.createCell(121).setCellValue(IDCardSelectePreview);
 	    row1.createCell(122).setCellValue(SelectedIDCardNPreview);
 	    row1.createCell(123).setCellValue(Identitimark01Preview);
 	    row1.createCell(124).setCellValue(Identitimark02Preview);
 	    
 	   WebElement Submit = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[text()='Submit']")));  
	    Submit.click();
	    Thread.sleep(1000);
	    Alert a = driver.switchTo().alert();
	    a.accept();
	    Thread.sleep(1000);
	    
	    switchToNewWindow(driver);
	    
	    WebElement forgot = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a[contains(text(),'Forget Application Number?')]")));
       act.moveToElement(forgot).click().perform();
	    
       switchToNewWindow(driver);
       
       WebElement adhar1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@name='aadharNumber']")));
       adhar1.sendKeys(Applicant_AadharNo);

       // Enter Date of Birth
       WebElement dateofbirth1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@name='dateofbirth']")));
       ((JavascriptExecutor) driver).executeScript("arguments[0].click();", dateofbirth1); 
       Dob(date, driver, "//input[@class='numInput cur-year']", "//select[contains(@class,'flatpickr-monthDropdown-months')]");
    
       WebElement login = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[ contains(text(),'Submit')]")));
       act.moveToElement(login).click().perform();
       Thread.sleep(2000);	
       jss.executeScript("window.scrollBy(0,1000)");

       // Wait for application number to appear
       WebElement appno = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("(//tr[@ class='odd' or @class='even']/td[1])[last()]")));
       //jss.executeScript("window.scrollBy(0,2000)");
       
       String appliationno = appno.getText();
       row1.createCell(125).setCellValue(appliationno);
       Thread.sleep(1000);	
       driver.findElement(By.xpath("(//button[contains(text(),'Close')])[2]")).click();
       Thread.sleep(1000);
       switchToNewWindow(driver);
       Thread.sleep(1000);
       WebElement myaap1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("(//button[contains(text(),'My Application')])[2]")));
       act.moveToElement(myaap1).click().perform();
       Thread.sleep(1000);
       switchToNewWindow(driver);
       
       WebElement apno = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ApplicantModel_ApplicationNo")));
       apno.sendKeys(appliationno);
       Thread.sleep(1000);
       WebElement dbo = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='ApplicantModel_DateOfBirth']")));
       ((JavascriptExecutor) driver).executeScript("arguments[0].click();", dbo); 
       Dob(date, driver, "//input[@class='numInput cur-year']", "//select[contains(@class,'flatpickr-monthDropdown-months')]");
    

       // Click 'Login'
       WebElement log = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[contains(text(),'Login')]")));
       act.moveToElement(log).click().perform();
       r.keyPress(KeyEvent.VK_ENTER);
       r.keyRelease(KeyEvent.VK_ENTER);	
       Thread.sleep(3000);
       jss.executeScript("window.scrollBy(0,500)");
       
       Thread.sleep(1000);
       WebElement download = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[text()='DOWNLOAD APPLICATION']")));
       Thread.sleep(2000);
       act.moveToElement(download).click().perform();
      Thread.sleep(1000);
      
      
       driver.findElement(By.linkText("Logout")).click();	
       switchToNewWindow(driver);
       
   	wait.until(ExpectedConditions.presenceOfElementLocated(By.linkText("New Application"))).click();
   	Thread.sleep(1000);
   	switchToNewWindow(driver);
   	jss.executeScript("window.scrollBy(0,1900)", "");
   	Thread.sleep(500);
   	
   	wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[1]"))).click();
   	Thread.sleep(500);
   	wait.until(ExpectedConditions.presenceOfElementLocated(By.id("nextBtn"))).click();
   	
   	switchToNewWindow(driver);
       
   	jss.executeScript("window.scrollBy(0,100)", "");
       
  //C://Users//pallavi//eclipse-workspace//project//TestData (2).xlsx
    Reporter.log(i +" iteration succesfully completed");
    System.out.println("ITERATION:");
    System.out.println(i +" iteration succesfully completed "); 
		}  catch (Exception e) {
	
	    	System.out.println("Failed: Error occurred in iteration " + i );
	        Reporter.log("Failed");
	        Reporter.log(i +" iteration is Skipping due to an error.");
	        Thread.sleep(1000);
	        String failed = i+" Failed";
	        row1.createCell(1).setCellValue(failed);
	       try {
	    	   WebElement Close = driver.findElement(By.xpath("(//button[contains(text(),'Close')])[2]"));
	       
	        if(isElementClickable(driver, Close)) {
	        Close.click();
	        }}
	       catch(Exception f) {
	    	 f.printStackTrace();
	       }		
	       
	        switchToNewWindow(driver);
	        wait.until(ExpectedConditions.presenceOfElementLocated(By.linkText("New Application"))).click();
	      	Thread.sleep(1000);
	      	switchToNewWindow(driver);
	      	jss.executeScript("window.scrollBy(0,1900)", "");
	      	Thread.sleep(500);
	      	
	    	Thread.sleep(500);
	      	wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[1]"))).click();
	      	Thread.sleep(1000);
	      	wait.until(ExpectedConditions.presenceOfElementLocated(By.id("nextBtn"))).click();
	      	
	      	switchToNewWindow(driver);
	          
	      	jss.executeScript("window.scrollBy(0,100)", "");
	        continue; 
	    	  
		} 
           
           fileOut = new FileOutputStream("D://Automation_data//PSI_Sports.xlsx");
   	    workbook.write(fileOut);
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
     driver.quit(); // Close the WebDriver session
      
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

     