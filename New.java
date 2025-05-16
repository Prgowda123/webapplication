package apc_kk;

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

public class New {


@Test
public void sample1() throws InterruptedException, AWTException, ParseException
{
	ChromeDriver driver = new ChromeDriver();
	driver.manage().window().maximize();
	driver.get("http://172.10.1.159:9016");
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
		FileInputStream fis1 = new FileInputStream("D://Automation_data//TestDataAPC_KK.xlsx");//"D:\steno\TestDataAPC.xlsx"
		XSSFWorkbook workbook = new XSSFWorkbook(fis1);
		Sheet sheet = workbook.getSheetAt(0);

		// Select select = null;
		int rowCount = sheet.getPhysicalNumberOfRows();

		// Loop through rows in the Excel sheet
		// int rowCount = sheet.getPhysicalNumberOfRows();




		for (int i =89; i <= 100; i++) { // Start from row 1 to skip header
		// Start from row 1 to skip header
			Row row = sheet.getRow(i);

			if (row == null) {
				System.out.println("Skipping empty row: " + i);
				continue;
			}

			if (row != null) {

				int applyingpost = (int) row.getCell(1).getNumericCellValue();
				int PostUnitCode = (int) row.getCell(2).getNumericCellValue();

				String Applicant_FullNameE = getCellValue(row.	getCell(3));
				String Fathername = getCellValue(row.getCell(4));
				String Applicant_MotherName = getCellValue(row.getCell(5));
				String Applicant_EmailId = getCellValue(row.getCell(6));
				String Applicant_MobileNo = getCellValue(row.getCell(7));
				String Applicant_AadharNo = getCellValue(row.getCell(8));
				String Applicant_DateOfBirth = getCellValue(row.getCell(9));
				int GenderCode =  (int) row.getCell(11).getNumericCellValue();
				String PhysicalDetail_Height = getCellValue(row.getCell(12));
				String DoorNo = getCellValue(row.getCell(13));
				String ContactAddress_Street = getCellValue(row.getCell(14));
				String ContactAddress_Taluk = getCellValue(row.getCell(15));
				String ContactAddress_City = getCellValue(row.getCell(16));
				int UnionStateCode =  (int) row.getCell(17).getNumericCellValue();
				int DistrictCode =  (int) row.getCell(18).getNumericCellValue();
				String Address_Pincode = getCellValue(row.getCell(20));
				String Address_Landmark = getCellValue(row.getCell(21));
				String NativeDistrict = getCellValue(row.getCell(22));
				String PermanentDoorNo = getCellValue(row.getCell(24));
				String PermanentAddress_Street = getCellValue(row.getCell(25));
				String PermanentAddress_Taluk = getCellValue(row.getCell(26));
				String PermanentAddress_City = getCellValue(row.getCell(27));
				int PermanentUnionStateCode =  (int) row.getCell(28).getNumericCellValue();
				int PermanentDistrictCode =  (int) row.getCell(29).getNumericCellValue();
				String Permanent_Pincode = getCellValue(row.getCell(31));
				String Permanent_Landmark = getCellValue(row.getCell(32));
				int CategoryCode =  (int) row.getCell(33).getNumericCellValue();
				String Reservation_SubCaste = getCellValue(row.getCell(34));
				String CategoryCertificateIssuedDate = getCellValue(row.getCell(35));
				String ExService_DischargeDate = getCellValue(row.getCell(45));
				int Exservicecatagorye =  (int) row.getCell(46).getNumericCellValue();
				int ExServiceForceCode =  (int) row.getCell(47).getNumericCellValue();
				int YearsInService =  (int) row.getCell(49).getNumericCellValue();
				int MonthsInService =  (int) row.getCell(50).getNumericCellValue();
				int DaysInService =  (int) row.getCell(51).getNumericCellValue();
				int ExServicemenRelationCode =  (int) row.getCell(41).getNumericCellValue();
				String Detail_JoiningDate = getCellValue(row.getCell(57));
				int YearsInService_govt =  (int) row.getCell(59).getNumericCellValue();
				int MonthsInService_govt =  (int) row.getCell(60).getNumericCellValue();
				int DaysInService_govt =  (int) row.getCell(61).getNumericCellValue();
				String GovermentServiceDetail_Department = getCellValue(row.getCell(62));
				String GovermentServiceDetail_Designation = getCellValue(row.getCell(63));
				String DepartmentEnquiryDetail = getCellValue(row.getCell(65));
				String CaseDetail = getCellValue(row.getCell(67));
				String ConvictionDetail = getCellValue(row.getCell(69));
				int QualificationBoardCode =  (int) row.getCell(71).getNumericCellValue();
				String otherboard = getCellValue(row.getCell(72));
				int KannadaLanguagePaper =  (int) row.getCell(73).getNumericCellValue();
				int passingyear =  (int) row.getCell(74).getNumericCellValue();
				String m_mark = getCellValue(row.getCell(76));
				String ob_mark = getCellValue(row.getCell(77));
				String gradess = getCellValue(row.getCell(79));
				String perss = getCellValue(row.getCell(80));

				String RegistrationNo = getCellValue(row.getCell(81));
				int QualificationBoardCodepu =  (int) row.getCell(83).getNumericCellValue();
				String OtherBoardName = getCellValue(row.getCell(84));
				int pucPassingyear =  (int) row.getCell(85).getNumericCellValue();
				String  pu_max = getCellValue(row.getCell(87));
				String pu_ob = getCellValue(row.getCell(88));
				String  Grade_Grade = getCellValue(row.getCell(90));
				String ScorePercentage = getCellValue(row.getCell(91));
				String puRegistrationNo = getCellValue(row.getCell(92));
				int IdentityCardTypeCode =  (int) row.getCell(94).getNumericCellValue();
				String  UploadedIDNo = getCellValue(row.getCell(95));
				String IdentificationMark_01 = getCellValue(row.getCell(96));
				String  IdentificationMark_02 = getCellValue(row.getCell(97));
				String Applicant_Photo = getCellValue(row.getCell(98));
				String  Applicant_Signature = getCellValue(row.getCell(99));
				String Applicant_Thumb = getCellValue(row.getCell(100));
				String  Applicant_IdentityCard = getCellValue(row.getCell(101));

				 String inputDate = Applicant_DateOfBirth; // Original format (dd-MMM-yyyy)

			        // Define input and output formats
			        SimpleDateFormat inputFormat = new SimpleDateFormat("dd-MMM-yyyy");
			        SimpleDateFormat outputFormat = new SimpleDateFormat("dd-MM-yyyy");

			        // Parse the input date string into a Date object
			        Date date = inputFormat.parse(inputDate);

			        // Format the Date object into the desired format
			        String formattedDate = outputFormat.format(date);

			        // Print the result
			        System.out.println("Converted Date: " + formattedDate); // Output: 21-03-2021

	WebElement ApplyingType = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplyingTypeCode")));
	Select s=new Select(ApplyingType);
	s.selectByIndex(applyingpost);

	WebElement unit = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PostUnitCode")));
	Select s1=new Select(unit);
	s1.selectByIndex(PostUnitCode);

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

	WebElement Dob = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_DateOfBirth']/following-sibling::input[1]")));
	Dob.sendKeys(Applicant_DateOfBirth);


	WebElement Gender = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GenderCode")));
	Select s2=new Select(Gender);
	s2.selectByIndex(GenderCode);

	WebElement Height = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_PhysicalDetail_Height")));

	Height.sendKeys(PhysicalDetail_Height);


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
	WebElement DateofSubCaste = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_CategoryCertificateIssuedDate']/following-sibling::input[1]")));
	DateofSubCaste.sendKeys(CategoryCertificateIssuedDate);

	WebElement Kannada = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsClamingKannadaMediumReservation' and @ value='True']")));
	act.moveToElement(Kannada).click().perform();

	WebElement PDP = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsClaimingPDPReservation' and @ value='False']")));
	act.moveToElement(PDP).click().perform();

	if(applyingpost == 2) {

		WebElement Armed = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_applicantExService_IsPresentlyServing' and @ value='True']")));
		act.moveToElement(Armed).click().perform();

		WebElement Benifit = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_applicantExService_IsAvaliedExServiceBenefit' and @ value='False']")));
		act.moveToElement(Benifit).click().perform();
		jss.executeScript("window.scrollBy(0,300)", "");

		WebElement Dateofdischarge = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='ApplicantExDischarge']/preceding-sibling::input[1]")));
		Dateofdischarge.sendKeys(ExService_DischargeDate);

		WebElement Exservicecatagory = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_applicantExService_ExServiceEducationalQualificationCode")));
		Select s8=new Select(Exservicecatagory);
		s8.selectByIndex(Exservicecatagorye);

		WebElement rendered  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_applicantExService_ExServiceForceCode")));
		Select s9=new Select(rendered );
		s9.selectByIndex(ExServiceForceCode);

		WebElement Exser_year  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_applicantExService_YearsInService")));
		Select s10=new Select(Exser_year );
		s10.selectByIndex(YearsInService);

		WebElement Exser_month  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_applicantExService_MonthsInService")));
		Select s11=new Select(Exser_month );
		s11.selectByIndex(MonthsInService);

		WebElement Exser_day  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_applicantExService_DaysInService")));
		Select s12=new Select(Exser_day );
		s12.selectByIndex(DaysInService);

		}

	else {
		WebElement Belong = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsClaimingExServicemenRelationReservation' and @ value='True']")));
		act.moveToElement(Belong).click().perform();

		WebElement Death = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_applicantExServiceFamily_IsDiedInService' and @ value='True']")));
		act.moveToElement(Death).click().perform();

		WebElement disable = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_applicantExServiceFamily_IsDisabledInService' and @ value='True']")));
		act.moveToElement(disable).click().perform();

		WebElement Relationship  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_applicantExServiceFamily_ExServicemenRelationCode")));
		Select s13=new Select(Relationship );
		s13.selectByIndex(ExServicemenRelationCode);
	}

	if(GenderCode ==2)
	{
		WebElement Transgender = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsClamingTransgenderReservation' and @ value='True']")));
		act.moveToElement(Transgender).click().perform();	
	}

	WebElement Rural = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsClaimingRuralReservation' and @ value='True']")));
	act.moveToElement(Rural).click().perform();

	if(applyingpost == 2) {
	WebElement KKJ = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsClamingKalyanaKarnatakaReservation' and @ value='False']")));
	act.moveToElement(KKJ).click().perform();

	WebElement KK = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsApplyingKKCertificate' and @ value='True']")));
	act.moveToElement(KK).click().perform();
	}

	else {
		WebElement KKJ = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsClamingKalyanaKarnatakaReservation' and @ value='True']")));
		act.moveToElement(KKJ).click().perform();
	}
	jss.executeScript("window.scrollBy(0,300)", "");

	if(applyingpost == 1 ) {
	WebElement GovtEmp = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_AreYouAGovermentEmployee' and @ value='True']")));
	act.moveToElement(GovtEmp).click().perform();

	WebElement Dateofjoining = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_GovermentServiceDetail_JoiningDate']/following-sibling::input[1]")));
	Dateofjoining.sendKeys(Detail_JoiningDate);

	WebElement ser_year  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GovermentServiceDetail_YearsInService")));
	Select s14=new Select(ser_year );
	s14.selectByIndex(YearsInService_govt);

	WebElement ser_month  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GovermentServiceDetail_MonthsInService")));
	Select s15=new Select(ser_month );
	s15.selectByIndex(MonthsInService_govt);

	WebElement ser_day  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GovermentServiceDetail_DaysInService")));
	Select s16=new Select(ser_day );
	s16.selectByIndex(DaysInService_govt);
	s16.selectByIndex(DaysInService_govt);	

	WebElement GovtDept = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GovermentServiceDetail_Department")));
	GovtDept.sendKeys(GovermentServiceDetail_Department);

	WebElement Designation = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GovermentServiceDetail_Designation")));
	Designation.sendKeys(GovermentServiceDetail_Designation);

	WebElement DeptEnquiry = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_CriminalActivity_HasDepartmentEnquiry' and @ value='True']")));
	act.moveToElement(DeptEnquiry).click().perform();

	WebElement DeptEnquirydetails = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_CriminalActivity_DepartmentEnquiryDetail")));
	DeptEnquirydetails.sendKeys(DepartmentEnquiryDetail);
	}

	else
	{
		WebElement GovtEmp = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_AreYouAGovermentEmployee' and @ value='False']")));
		act.moveToElement(GovtEmp).click().perform();

	}

	WebElement Criminalcase = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_CriminalActivity_IsInvolvedInCriminalActivity' and @ value='True']")));
	act.moveToElement(Criminalcase).click().perform();

	WebElement Criminalcasedetails = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_CriminalActivity_CaseDetail")));
	Criminalcasedetails.sendKeys(CaseDetail);

	WebElement ConvictedInCriminalCase = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_CriminalActivity_IsConvictedInCriminalCase' and @ value='True']")));
	act.moveToElement(ConvictedInCriminalCase).click().perform();

	WebElement ConvictedInCriminalCasedetails = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_CriminalActivity_ConvictionDetail")));
	ConvictedInCriminalCasedetails.sendKeys(ConvictionDetail);

	jss.executeScript("window.scrollBy(0,300)", "");

	//For Educational Qualification Details SSLC

	WebElement SSLC = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_IsSSLCHolder' and @ value='True']")));
	Thread.sleep(1500);
    act.moveToElement(SSLC).click().perform();

	jss.executeScript("window.scrollBy(0,300)", "");
	Thread.sleep(100);
	WebElement Board  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_QualificationBoardCode")));
	Select s17=new Select(Board );
	s17.selectByIndex(QualificationBoardCode);

	if(QualificationBoardCode==3)
	{
		WebElement OtherBoard = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_OtherBoardName")));
		OtherBoard.sendKeys(otherboard);
	}

	WebElement paper  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_KannadaLanguagePaper")));
	Select s18=new Select(paper );
	s18.selectByIndex(KannadaLanguagePaper);
	WebElement Passingyear  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_YearOfPassing")));
	Select s19=new Select(Passingyear );
	s19.selectByIndex(passingyear);

	if(applyingpost == 1)
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

		WebElement Puc = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_IsPUCHolder' and @ value='True']")));
		act.moveToElement(Puc).click().perform();

		jss.executeScript("window.scrollBy(0,300)", "");

		WebElement PUBoard  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_QualificationBoardCode")));
		Select s20=new Select(PUBoard );
		s20.selectByIndex(QualificationBoardCodepu);

		if(QualificationBoardCodepu==3)
		{
			WebElement PUOtherBoard = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_OtherBoardName")));
			PUOtherBoard.sendKeys(OtherBoardName);
		}
		WebElement PUPassingyear  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_YearOfPassing")));
		Select s21=new Select(PUPassingyear);
		s21.selectByIndex(pucPassingyear);

		if(applyingpost==1) {
		WebElement PUMarkorgrade = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_PUCQualification_MarkType' and @ value='G']")));
		act.moveToElement(PUMarkorgrade).click().perform();

		WebElement Grade = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_Grade_Grade")));
		Grade.sendKeys(Grade_Grade);

		WebElement Per = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_ScorePercentage")));
		Per.sendKeys(ScorePercentage);
			}

		else
		{
			WebElement PUMarkorgrade = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_PUCQualification_MarkType' and @ value='M']")));
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

		WebElement Degree = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_IsDegreeHolder' and @ value='True']")));
		act.moveToElement(Degree).click().perform();
		jss.executeScript("window.scrollBy(0,400)", "");

		//For Documents Upload 
		WebElement IDcard = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_IdentityCardTypeCode")));
		Select s22=new Select(IDcard );
		s22.selectByIndex(IdentityCardTypeCode);


		if(IdentityCardTypeCode !=1) {
			WebElement IDcardno = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_UploadedIDNo")));
			IDcardno.sendKeys(UploadedIDNo);
		}

		WebElement Mark1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_IdentificationMark_01")));
		Mark1.sendKeys(IdentificationMark_01);

		WebElement Mark2 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_IdentificationMark_02")));
		Mark2.sendKeys(IdentificationMark_02);


	       WebElement photo=wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Photo")));
	       photo.sendKeys(Applicant_Photo);//

	       WebElement sign=wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Signature")));
	       sign.sendKeys(Applicant_Signature);

	       WebElement thumb=wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Thumb")));
	       thumb.sendKeys(Applicant_Thumb);

	       WebElement ID=wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_IdentityCard")));
	       ID.sendKeys(Applicant_IdentityCard);

	       WebElement preview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("preview-btn")));
	       preview.click();
	       r.keyPress(KeyEvent.VK_ENTER);
		   r.keyRelease(KeyEvent.VK_ENTER);


	       //Preview contents
	       Sheet sheet2 = workbook.getSheetAt(3);
		   Row row1 = sheet2.createRow(sheet2.getPhysicalNumberOfRows()); 
		   WebElement appdetails = driver.findElement(By.xpath("//h5[@id='exampleModalCenterTitle']"));
			String details1 = appdetails.getText();

			if(!isElementClickable(driver, appdetails)) {try {
			    boolean hasError = true;

			    // List of specific error IDs (extend this as needed)
			    String[] errorIds = {"Applicant_ApplyingTypeCode-error","Applicant_PostUnitCode-error",
			        "Applicant_FullName-error", "Applicant_FatherName-error", "Applicant_MotherName-error",
			        "Applicant_EmailId-error", "Applicant_MobileNo-error", "Applicant_AadharNo-error", "Applicant_DateOfBirth-error",
			        "Applicant_Reservation_GenderCode-error","Applicant_PhysicalDetail_Height-error","Applicant_ContactAddress_DoorNo-error","Applicant_ContactAddress_Street-error",
			        "Applicant_ContactAddress_Taluk-error","Applicant_ContactAddress_OtherDistrictName-error","Applicant_ContactAddress_Pincode",
			        "Applicant_ContactAddress_Landmark-error","Applicant_NativeDistrict-error","Applicant_PermanentAddress_DoorNo-error",
			        "Applicant_PermanentAddress_Street-error","Applicant_PermanentAddress_Taluk-error","Applicant_PermanentAddress_City-error",
			        "Applicant_PermanentAddress_OtherDistrictName-error","Applicant_PermanentAddress_Pincode","Applicant_PermanentAddress_Landmark-error",
			      "Applicant_Reservation_CategoryCode-error","Applicant_Reservation_SubCaste-error","Applicant_Reservation_GovermentServiceDetail_YearsInService-error","Applicant_Reservation_GovermentServiceDetail_Department-error",
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
			}}


	    try {   WebElement candidateTypePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@id='candidateTypePreview']")));
	       wait.until(ExpectedConditions.visibilityOf(candidateTypePreview));
			wait.until(ExpectedConditions.elementToBeClickable(candidateTypePreview));
	       String candidatePreview = candidateTypePreview.getText();


	       WebElement unitNamePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("unitNamePreview")));
	       wait.until(ExpectedConditions.visibilityOf(unitNamePreview));
			wait.until(ExpectedConditions.elementToBeClickable(unitNamePreview));
	       String unitPreview = unitNamePreview.getText();

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
	       jss.executeScript("arguments[0].scrollIntoView(true);", HeightPreview);	

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

	       //reservation
	       WebElement CategoryPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("CategoryPreview")));
	       String CategoriPreview = CategoryPreview.getText();


	       WebElement SubcastePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SubcastePreview")));
	       String SubcasteePreview = SubcastePreview.getText();

	       WebElement DateofSubcastePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DateofSubcastePreview")));
	       String DateofSubcasteePreview = DateofSubcastePreview.getText();
	       jss.executeScript("arguments[0].scrollIntoView(true);", DateofSubcastePreview);

	       WebElement KannadaMediumReservationPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("KannadaMediumReservationPreview")));
	       String KannadaMediummReservationPreview = KannadaMediumReservationPreview.getText();

	       WebElement PDPReservationPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PDPReservationPreview")));
	       String PDPReservationnPreview = PDPReservationPreview.getText();

	       row1.createCell(1).setCellValue(candidatePreview);
	       row1.createCell(2).setCellValue(unitPreview);
	       row1.createCell(3).setCellValue(candidatenamePreview);
	       row1.createCell(4).setCellValue(fatherPreview);
	       row1.createCell(5).setCellValue(MotherPreview);
	       row1.createCell(6).setCellValue(emailidPreview);
	       row1.createCell(7).setCellValue(MobilePreview);
	       row1.createCell(8).setCellValue(aadharnoPreview);
	       row1.createCell(9).setCellValue(DateofBirtPreview);
           row1.createCell(10).setCellValue(DateofBirthasonnPreview);
           row1.createCell(11).setCellValue(genderrPreview);
           row1.createCell(12).setCellValue(HeighttPreview);
           row1.createCell(13).setCellValue(DoornoPreview);
           row1.createCell(14).setCellValue(StreettPreview);
           row1.createCell(15).setCellValue(talukkPreview);
           row1.createCell(16).setCellValue(citiPreview);
           row1.createCell(17).setCellValue(stateePreview);
           row1.createCell(18).setCellValue(districttPreview);
           row1.createCell(20).setCellValue(pincodeePreview);
           row1.createCell(21).setCellValue(landmarkkPreview);
           row1.createCell(22).setCellValue(nativeDistrictP);
           row1.createCell(23).setCellValue(postaladdresPreview);
           row1.createCell(24).setCellValue(PermDoorNoPreview);
           row1.createCell(25).setCellValue(PermStreetPreview);
           row1.createCell(26).setCellValue(PermTalukPreview);
           row1.createCell(27).setCellValue(PermCityPreview);
           row1.createCell(28).setCellValue(permstatePreview);
           row1.createCell(29).setCellValue(PermDistrictPreview);
           row1.createCell(31).setCellValue(PermPincodePreview);
           row1.createCell(32).setCellValue(NearbyLandmarkkPreview);
           row1.createCell(33).setCellValue(CategoriPreview);
           row1.createCell(34).setCellValue(SubcasteePreview);
           row1.createCell(35).setCellValue(DateofSubcasteePreview);
           row1.createCell(36).setCellValue(KannadaMediummReservationPreview);
           row1.createCell(37).setCellValue(PDPReservationnPreview);



	       if(applyingpost !=1)
	       {

	    	   WebElement exServicemenPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("exServicemenPreview")));
		       String exServicemennPreview = exServicemenPreview.getText();

		       WebElement presentlyExServicemenPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("presentlyExServicemenPreview")));
		       String presentlyyExServicemenPreview = presentlyExServicemenPreview.getText();

		       WebElement availedBenefitPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("availedBenefitPreview")));
		       String availedBenefittPreview = availedBenefitPreview.getText();

		       WebElement ExServicemendateofdisexserPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ExServicemendateofdisexserPreview")));
		       String ExServicemendateeofdisexserPreview = ExServicemendateofdisexserPreview.getText();


		       WebElement eduQualificationexserPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("eduQualificationexserPreview")));
		       String eduQualificationexserrPreview = eduQualificationexserPreview.getText();

		       WebElement ExServicemenServiceRenderedPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ExServicemenServiceRenderedPreview")));
		       String ExServicemenServiceRendereddPreview = ExServicemenServiceRenderedPreview.getText();
		       jss.executeScript("arguments[0].scrollIntoView(true);", ExServicemenServiceRenderedPreview);

		       WebElement yearsOfExServicemenPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("yearsOfExServicemenPreview")));
		       String yearsOfExServicemennPreview = yearsOfExServicemenPreview.getText();

		       WebElement MonthsOfExServicemenPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("MonthsOfExServicemenPreview")));
		       String MonthsOfExServicemennPreview = MonthsOfExServicemenPreview.getText();


		       WebElement DaysOfExServicemenPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DaysOfExServicemenPreview")));
		       String DaysOfExServicemennPreview = DaysOfExServicemenPreview.getText();

		       row1.createCell(42).setCellValue(exServicemennPreview);
		       row1.createCell(43).setCellValue(presentlyyExServicemenPreview);
		       row1.createCell(44).setCellValue(availedBenefittPreview);
		       row1.createCell(45).setCellValue(ExServicemendateeofdisexserPreview);
		       row1.createCell(46).setCellValue(eduQualificationexserrPreview);
		       row1.createCell(47).setCellValue(ExServicemenServiceRendereddPreview);
		       row1.createCell(49).setCellValue(yearsOfExServicemennPreview);
		       row1.createCell(50).setCellValue(MonthsOfExServicemennPreview);
		       row1.createCell(51).setCellValue(DaysOfExServicemennPreview);

	       }
	       else {
	    	   WebElement RelationShipwithExServicemenPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("RelationShipwithExServicemenPreview")));
		       String RelationShipwithExServicemennPreview = RelationShipwithExServicemenPreview.getText();

		       WebElement ExServicemenDiedinActionPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ExServicemenDiedinActionPreview")));
		       String ExServicemenDiedinActionnPreview = ExServicemenDiedinActionPreview.getText();
		       jss.executeScript("arguments[0].scrollIntoView(true);", ExServicemenDiedinActionPreview);  

		       WebElement ExServicemenDisabledPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ExServicemenDisabledPreview")));
		       String ExServicemenDisablePreview = ExServicemenDisabledPreview.getText();

		       WebElement ExServicemenRelationPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ExServicemenRelationPreview")));
		       String ExServicemenRelationnPreview = ExServicemenRelationPreview.getText();
		       jss.executeScript("arguments[0].scrollIntoView(true);", ExServicemenRelationPreview);  
		       row1.createCell(38).setCellValue(RelationShipwithExServicemennPreview);
		       row1.createCell(39).setCellValue(ExServicemenDiedinActionnPreview);
		       row1.createCell(40).setCellValue(ExServicemenDisablePreview);
		       row1.createCell(41).setCellValue(ExServicemenRelationnPreview);

	       }  	

	       if(GenderCode ==2) {
	    	   WebElement TransgenderReservationPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("TransgenderReservationPreview")));
		       String TransgenderReservationnPreview = TransgenderReservationPreview.getText();
		      row1.createCell(52).setCellValue(TransgenderReservationnPreview);
	       }

	       WebElement ClaimingRuralMediumPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ClaimingRuralMediumPreview")));
	       String ClaimingRuralMediummPreview = ClaimingRuralMediumPreview.getText();
	       jss.executeScript("arguments[0].scrollIntoView(true);", ClaimingRuralMediumPreview);  
           row1.createCell(53).setCellValue(ClaimingRuralMediummPreview); 

	       if(applyingpost == 2) {
	       WebElement KalyanaKarnatakaPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("KalyanaKarnatakaPreview")));
	       String KalyanaKarnatakaaPreview = KalyanaKarnatakaPreview.getText();

	       WebElement KalyanaKarnatakaDistrictPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("KalyanaKarnatakaDistrictPreview")));
	       String KalyanaKarnatakaaDistrictPreview = KalyanaKarnatakaDistrictPreview.getText();
	       row1.createCell(54).setCellValue(KalyanaKarnatakaaPreview);
	       row1.createCell(55).setCellValue(KalyanaKarnatakaaDistrictPreview);
	       }
	       else {
	    	   WebElement KalyanaKarnatakaPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("KalyanaKarnatakaPreview")));
		       String KalyanaKarnatakaaPreview = KalyanaKarnatakaPreview.getText();
		       row1.createCell(54).setCellValue(KalyanaKarnatakaaPreview);
	       }

	       if(applyingpost == 1 ) {
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

	       WebElement GovtDepartmentPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("GovtDepartmentPreview")));
	       String GovtnDepartmentPreview = GovtDepartmentPreview.getText();

	       WebElement GovtDesignationPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("GovtDesignationPreview")));
	       String GovtnDesignationPreview = GovtDesignationPreview.getText();

	       WebElement DepartmentalEnquirPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DepartmentalEnquirPreview")));
	       String DepartmentalEnquiryPreview = DepartmentalEnquirPreview.getText();

	       WebElement DeptenqdetailsPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DeptenqdetailsPreview")));
	       String DeptenqrydetailsPreview = DeptenqdetailsPreview.getText();

	       row1.createCell(56).setCellValue(GovernmentEmployePreview);
	       row1.createCell(57).setCellValue(GovtDateofJoininggPreview);
	       row1.createCell(59).setCellValue(GovtYearsPrevieww);
	       row1.createCell(60).setCellValue(GovtMonthhPreview);
	       row1.createCell(61).setCellValue(GovtDayPreview);
	       row1.createCell(62).setCellValue(GovtnDepartmentPreview);
	       row1.createCell(63).setCellValue(GovtnDesignationPreview);
	       row1.createCell(64).setCellValue(DepartmentalEnquiryPreview);
	       row1.createCell(65).setCellValue(DeptenqrydetailsPreview);

	       }
	       else
	       {
	    	  WebElement GovernmentEmployeePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("GovernmentEmployeePreview")));
	  	      String GovernmentEmployePreview = GovernmentEmployeePreview.getText();
	  	    row1.createCell(56).setCellValue(GovernmentEmployePreview);
	       }

	       WebElement CriminalCasesPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("CriminalCasesPreview")));
	       String CriminalCasePreview = CriminalCasesPreview.getText();   

	       WebElement CriminalCasesdetailsPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("CriminalCasesdetailsPreview")));
	  	   String CriminalCasesdetailPreview = CriminalCasesdetailsPreview.getText();

	  	   WebElement ConvictedinaCriminalPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ConvictedinaCriminalPreview")));
 	       String ConvictedinaCriminalsPreview = ConvictedinaCriminalPreview.getText();

 	      WebElement ConvictedCriminalDetailsPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ConvictedCriminalDetailsPreview")));
 	      String ConvictedCriminalDetailPreview = ConvictedCriminalDetailsPreview.getText(); 
	      jss.executeScript("arguments[0].scrollIntoView(true);", ConvictedCriminalDetailsPreview); 

	       row1.createCell(66).setCellValue(CriminalCasePreview);
	       row1.createCell(67).setCellValue(CriminalCasesdetailPreview);
	       row1.createCell(68).setCellValue(ConvictedinaCriminalsPreview);
	       row1.createCell(69).setCellValue(ConvictedCriminalDetailPreview);



	     //sslc  

 	     WebElement PassedSSLCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PassedSSLCPreview")));
	     String PasseSSLCPreview = PassedSSLCPreview.getText(); 

	     WebElement BoardofSslcPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("BoardofSslcPreview")));
	     String BoardSslcPreview = BoardofSslcPreview.getText(); 

	     row1.createCell(70).setCellValue(PasseSSLCPreview);
	     row1.createCell(71).setCellValue(BoardSslcPreview);


	      if(QualificationBoardCode==3)
	      {
	     WebElement OtherSslcBoarPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("OtherSslcBoarPreview")));
	     String OtherSslcBoardPreview = OtherSslcBoarPreview.getText(); 
	     row1.createCell(72).setCellValue(OtherSslcBoardPreview);

	      }  
	     WebElement KannadaLanguagePreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("KannadaLanguagePreview")));
	     String KannadaLanguagPreview = KannadaLanguagePreview.getText(); 


	     WebElement YearofPassingSSLCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("YearofPassingSSLCPreview")));
	     String YearfPassingSSLCPreview = YearofPassingSSLCPreview.getText(); 

	     WebElement SSLCMarksorGradesPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SSLCMarksorGradesPreview")));
	     String SSLCMarksorGradePreview = SSLCMarksorGradesPreview.getText(); 
	     if(applyingpost ==1)  { 
	     WebElement SSLCMaxMarksPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SSLCMaxMarksPreview")));
	     String SSLCMaxMarkPreview = SSLCMaxMarksPreview.getText(); 
	      jss.executeScript("arguments[0].scrollIntoView(true);", SSLCMaxMarksPreview);   


	     WebElement SSLCMarksObtainedPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SSLCMarksObtainedPreview")));
	     String SSLCMarkObtainedPreview = SSLCMarksObtainedPreview.getText(); 
	     row1.createCell(76).setCellValue(SSLCMaxMarkPreview);
	     row1.createCell(77).setCellValue(SSLCMarkObtainedPreview);

	     WebElement SSLCPercentageObtainedPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SSLCPercentageObtainedPreview")));
	     String SSLCPercentageObtainePreview = SSLCPercentageObtainedPreview.getText(); 
	     row1.createCell(78).setCellValue(SSLCPercentageObtainePreview);
	     }  

	     else {
	    	   WebElement SSLCGradesObtainedPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SSLCGradesObtainedPreview")));
	  	     String SSLCGradesObtainePreview = SSLCGradesObtainedPreview.getText(); 

	  	   WebElement SSLCPercentageObtainedPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SSLCPercentageObtainedPreview")));
		     String SSLCPercentageObtainePreview = SSLCPercentageObtainedPreview.getText(); 
		     row1.createCell(79).setCellValue(SSLCGradesObtainePreview);
		     row1.createCell(80).setCellValue(SSLCPercentageObtainePreview);

	     }
	     WebElement SSLCRegistrationNoPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("SSLCRegistrationNoPreview")));
	     String SSLCRegistrationsNoPreview = SSLCRegistrationNoPreview.getText(); 

	     row1.createCell(73).setCellValue(KannadaLanguagPreview);
	     row1.createCell(74).setCellValue(YearfPassingSSLCPreview);
	     row1.createCell(75).setCellValue(SSLCMarksorGradePreview);


	     row1.createCell(81).setCellValue(SSLCRegistrationsNoPreview);


	     //puc
	    WebElement PassedPUCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PassedPUCPreview")));
	    String PassePUCPreview = PassedPUCPreview.getText(); 
	    jss.executeScript("arguments[0].scrollIntoView(true);", PassedPUCPreview);  


	    WebElement PassedPucBoardPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PassedPucBoardPreview")));
	    String PassePucBoardPreview = PassedPucBoardPreview.getText(); 

	    row1.createCell(82).setCellValue(PassePUCPreview);
	    row1.createCell(83).setCellValue(PassePucBoardPreview);

	    if(QualificationBoardCodepu==3) {
	    WebElement PucOtherBoardPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PucOtherBoardPreview")));
	    String PuOtherBoardPreview = PucOtherBoardPreview.getText(); 
	    row1.createCell(84).setCellValue(PuOtherBoardPreview);

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
	    row1.createCell(90).setCellValue(GradeObtainedPUCPreview);
	    row1.createCell(91).setCellValue(PercentagePUPreview);
	    }
	    else {
		    WebElement MaxMarksPUCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("MaxMarksPUCPreview")));
		    String MaxMarksPUCPrevie = MaxMarksPUCPreview.getText(); 

		    WebElement MarksobtainedPUCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("MarksobtainedPUCPreview")));
		    String MarksobtainedPUCPrevie = MarksobtainedPUCPreview.getText(); 

	     WebElement PercentagePUCPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PercentagePUCPreview")));
	 	 String PercentagePUPreview = PercentagePUCPreview.getText(); 
	 	 row1.createCell(87).setCellValue(MaxMarksPUCPrevie);
	 	 row1.createCell(88).setCellValue(MarksobtainedPUCPrevie);
	 	 row1.createCell(89).setCellValue(PercentagePUPreview);
	    }
	    WebElement PUCRegistrationsnoPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("PUCRegistrationsnoPreview")));
	    String PUCRegistrationnoPreview = PUCRegistrationsnoPreview.getText(); 
	    jss.executeScript("arguments[0].scrollIntoView(true);", PUCRegistrationsnoPreview); 

	     row1.createCell(85).setCellValue(YearPassingPUCPreview);
	     row1.createCell(86).setCellValue(MarkorGradePUCPreview);

	     row1.createCell(92).setCellValue(PUCRegistrationnoPreview);


	    WebElement DegHolderPreview = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("DegHolderPreview")));
	    String DegreeHolderPreview = DegHolderPreview.getText(); 
	    row1.createCell(93).setCellValue(DegreeHolderPreview);

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
	    row1.createCell(94).setCellValue(IDCardSelectePreview);
	    row1.createCell(95).setCellValue(SelectedIDCardNPreview);
	    row1.createCell(96).setCellValue(Identitimark01Preview);
	    row1.createCell(97).setCellValue(Identitimark02Preview);

	    WebElement Submit = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[text()='Submit']")));  
	    Submit.click();
	    Thread.sleep(1000);
	    Alert a = driver.switchTo().alert();
	    a.accept();
	    Thread.sleep(1000);

	    switchToNewWindow(driver);

	    WebElement forgot = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a[contains(text(),'Forgot Application No?')]")));
        act.moveToElement(forgot).click().perform();

        switchToNewWindow(driver);

        WebElement adhar1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='aadharNumber']")));
        adhar1.sendKeys(Applicant_AadharNo);

        // Enter Date of Birth
        WebElement dateofbirth1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@class='form-control dob flatpickr-input']")));
        dateofbirth1.sendKeys(formattedDate	);

        // Click 'Submit'
        WebElement login = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[@id='submitBtn']")));
        act.moveToElement(login).click().perform();
        Thread.sleep(2000);	
        jss.executeScript("window.scrollBy(0,1000)");

        // Wait for application number to appear
        WebElement appno = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("(//tr[@ class='odd' or @class='even']/td[1])[last()]")));
        //jss.executeScript("window.scrollBy(0,2000)");

        String appliationno = appno.getText();
        row1.createCell(98).setCellValue(appliationno);
        Thread.sleep(2000);	
        driver.findElement(By.xpath("(//button[contains(text(),'Close')])[2]")).click();
        Thread.sleep(2000);

        switchToNewWindow(driver);
        Thread.sleep(1000);
        WebElement myaap1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("(//a[contains(text(),'My Application')])[2]")));
        act.moveToElement(myaap1).click().perform();
        Thread.sleep(1000);
        switchToNewWindow(driver);

        WebElement apno = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ApplicantModel_ApplicationNo")));
        apno.sendKeys(appliationno);
        Thread.sleep(1000);
        WebElement dbo = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ApplicantModel_DateOfBirth")));
        dbo.sendKeys(formattedDate);

        // Click 'Login'
        WebElement log = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[contains(text(),'Login')]")));
        act.moveToElement(log).click().perform();
        r.keyPress(KeyEvent.VK_ENTER);
        r.keyRelease(KeyEvent.VK_ENTER);	
        Thread.sleep(4000);
        jss.executeScript("window.scrollBy(0,500)");

        WebElement download = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[text()='DOWNLOAD APPLICATION']")));
        Thread.sleep(1000);
        act.moveToElement(download).click().perform();
        Thread.sleep(2000);
        driver.findElement(By.linkText("Logout")).click();	
        switchToNewWindow(driver);
    	wait.until(ExpectedConditions.presenceOfElementLocated(By.linkText("New Application"))).click();
    	jss.executeScript("window.scrollBy(0,1900)", "");
    	Thread.sleep(500);
    	switchToNewWindow(driver);
    	wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[1]"))).click();
    	Thread.sleep(500);
    	wait.until(ExpectedConditions.presenceOfElementLocated(By.id("nextBtn"))).click();

    	switchToNewWindow(driver);

    	jss.executeScript("window.scrollBy(0,100)", "");


	//C://Users//pallavi//eclipse-workspace//project//TestData (2).xlsx
    Reporter.log(i +" iteration succesfully completed");
    System.out.println("ITERATION:");
    System.out.println(i +" iteration succesfully completed ");
	    }
	    catch (Exception e) {

	    	System.out.println("Failed: Error occurred in iteration " + i );
	        Reporter.log("Failed");
	        Reporter.log(i +" iteration is Skipping due to an error.");
	        Thread.sleep(5000);
	        String failed = i+" Failed";
	        row1.createCell(1).setCellValue(failed);
	       try {
	    	   WebElement Close = driver.findElement(By.linkText("Close"));

	        if(isElementClickable(driver, Close)) {
	        Close.click();
	        }}
	       catch(Exception f) {

	       }

	        switchToNewWindow(driver);
	        wait.until(ExpectedConditions.presenceOfElementLocated(By.linkText("New Application"))).click();
	      	jss.executeScript("window.scrollBy(0,1900)", "");
	      	Thread.sleep(500);
	      	switchToNewWindow(driver);
	      	wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[1]"))).click();
	      	Thread.sleep(1000);
	      	wait.until(ExpectedConditions.presenceOfElementLocated(By.id("nextBtn"))).click();

	      	switchToNewWindow(driver);

	      	jss.executeScript("window.scrollBy(0,100)", "");
	        continue; 

		}

		fileOut = new FileOutputStream("D://Automation_data//TestDataAPC_KK.xlsx");
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
