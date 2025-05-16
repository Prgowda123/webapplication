package apc_nkk;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Sleeper;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class Title {
@Test
public void sample() throws AWTException, InterruptedException, IOException
{
	
	
	
	ChromeDriver driver = new ChromeDriver();
	driver.manage().window().maximize();
	driver.get("http://172.10.1.159:9017");
	WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
	wait.until(ExpectedConditions.presenceOfElementLocated(By.linkText("New Application"))).click();
	switchToNewWindow(driver);
	Robot r=new Robot();
	Actions act = new Actions(driver);
	JavascriptExecutor jss = (JavascriptExecutor) driver;
	jss.executeScript("window.scrollBy(0,1500)", "");
	wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[1]"))).click();
	wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[contains(text(),'Continue To Application')]"))).click();
    switchToNewWindow(driver);
    jss.executeScript("window.scrollBy(0,150)", "");
    
    WebElement Applying_as = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//label[contains(text(),'Applying as a ')]")));
    String Applying_as_a = Applying_as.getText();
    WebElement ApplyingType = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplyingTypeCode")));
	Select s=new Select(ApplyingType);
	s.selectByIndex(1);
	
    WebElement unit = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//label[contains(text(),'Unit Name ')]")));
    String unit_name = unit.getText();
    
    WebElement Applicant_Full = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Applicant Full Name')]")));
    String Applicant_Full_name = Applicant_Full.getText();
    
    WebElement Father = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Father')]")));
    String Father_name = Father.getText();
    
    WebElement Mother = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Mother')]")));
    String Mother_name = Mother.getText();
    
    WebElement Email = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Email ID')]")));
    String Email_id = Email.getText();
    
    WebElement Mobile = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Mobile Number')]")));
    String Mobile_no = Mobile.getText();
    
    WebElement Aadhar = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Aadhar Number')]")));
    String Aadhar_no = Aadhar.getText();
    
    WebElement Date = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Date of Birth')]")));
    String dob = Date.getText();
    
    WebElement age = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Age of Applicant as')]")));
    String age_applicant = age.getText();
    
    WebElement gender = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Gender')]")));
    String gen = gender.getText();
    
    WebElement Gender = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_GenderCode")));
	Select s2=new Select(Gender);
	s2.selectByIndex(1);
    
	WebElement height = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Height of the Applicant in cm')]")));
    String height_applicant = height.getText();
	
    //address
    
    WebElement dooor = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Door No')]")));
    String door_no = dooor.getText();
	
    WebElement stret = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Street ')]")));
    String street = stret.getText();
	
    WebElement Taluk = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Taluk ')]")));
    String taluk = Taluk.getText();
	
    WebElement City = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'City ')]")));
    String city = City.getText();
	
    WebElement Statee = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'State ')]")));
    String state = Statee.getText();
	
    WebElement State = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_UnionStateCode")));
	Select s3=new Select(State);
	Thread.sleep(1000);
	s3.selectByIndex(14);
	r.keyPress(KeyEvent.VK_TAB);
	r.keyRelease(KeyEvent.VK_TAB);
	
	WebElement District = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ContactAddress_DistrictCode")));
	Select s4=new Select(District);
	Thread.sleep(1000);
	s4.selectByIndex(32);
	
	WebElement dis = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'District')]")));
    String district = dis.getText();
	
    WebElement othdis = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'If selected other district mention the Specify District')]")));
    String other_district = othdis.getText();
	 
    WebElement pincode = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Pincode ')]")));
    String Pincode = pincode.getText();
    
    WebElement Landmark = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Nearby Landmark')]")));
    String landmark = Landmark.getText();
    
    WebElement Native = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Native District')]")));
    String Native_dis = Native.getText();
  	jss.executeScript("arguments[0].scrollIntoView(true);", Native);
  	 WebElement perm = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Postal Address and Permanent Address are Same')]")));
     String peradd = perm.getText();
    //Reservation
    
    WebElement Category = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Category ')]")));
    String category = Category.getText();
    
	WebElement catagory = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_Reservation_CategoryCode")));
	Select s7=new Select(catagory);
	s7.selectByIndex(3);
	
	WebElement Sub_Caste = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Sub Caste')]")));
    String SubCaste = Sub_Caste.getText();
    jss.executeScript("arguments[0].scrollIntoView(true);", Sub_Caste);

    WebElement CasteCertificate = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Date of Issue of  Caste Certificate')]")));
    String Caste_Certificate = CasteCertificate.getText();
    
    WebElement KannadaMedium = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Are you claiming Kannada Medium Reservation')]")));
    String Kannada_Medium = KannadaMedium.getText();
    
    WebElement PDP = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Are you claiming PDP Reservation')]")));
    String pdp = PDP.getText();
    
    WebElement belong_to_afamily = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Do you belong to a family of an Ex-Servicemen who while in service')]")));
    String belong_to_a_family = belong_to_afamily.getText();
    
	WebElement Belong = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_IsClaimingExServicemenRelationReservation' and @ value='True']")));
	act.moveToElement(Belong).click().perform();
	
    WebElement Deathwhileinservice = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Death while in service')]")));
    String Deathwhile_in_service = Deathwhileinservice.getText();
    jss.executeScript("arguments[0].scrollIntoView(true);", Deathwhileinservice);
    
    Thread.sleep(1000);
    

    WebElement disabled = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Permanently disabled while in Service')]")));
    String Permanentlydisabled = disabled.getText();
  
    WebElement Relationship = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Relationship with the Ex-Servicemen')]")));
    String Relationship_Exservice = Relationship.getText();
    
    Thread.sleep(2000);
    jss.executeScript("window.scrollBy(0,-2500)", "");
    
    WebElement ApplyingType1 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_ApplyingTypeCode")));
   	Select ss3=new Select(ApplyingType1);
   	ss3.selectByIndex(2);
   	
    Thread.sleep(2000);
    jss.executeScript("window.scrollBy(0,2500)", "");
    
    WebElement ExServicemen = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Are you an Ex-Servicemen')]")));
    String Ex_Servicemen = ExServicemen.getText();
     
    WebElement ArmedForce = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Are you presently serving in Armed Force')]")));
    String Armed_Force = ArmedForce.getText();
    
    WebElement benefit = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Have you availed Ex-Servicemen benefit')]")));
    String benefitexs = benefit.getText();
    jss.executeScript("arguments[0].scrollIntoView(true);", benefit);

    
    WebElement Discharge = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Date of Discharge from Service')]")));
    String Dischargedate = Discharge.getText();
     
    WebElement Educational = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Educational Qualification Ex-Servicemen')]")));
    String Educational_Qualification = Educational.getText();
    
    WebElement Service_rendered = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Service rendered in which force (Army/Air Force/Navy)')]")));
    String Servicerendered = Service_rendered.getText();
    
    WebElement Service = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'How many years of Service have you rendered')]")));
    String Serviceyear = Service.getText();
      
    WebElement year = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//label[contains(text(),'Years')]")));
    String years = year.getText();
     
    WebElement Months = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//label[contains(text(),'Months')]")));
    String Month = Months.getText();
     
    WebElement Days = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//label[contains(text(),'Days ')]")));
    String Day = Days.getText();    
  
    WebElement Transgender = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Are you claiming Transgender Reservation')]")));
    String Transgender_ser = Transgender.getText();
    
    WebElement Ruralmedium = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Are you claiming Rural medium Reservation')]")));
    String Rural_medium = Ruralmedium.getText();     
  
    
  
	   
    WebElement Government_Employee = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Are you a Government Employee')]")));
    String GovernmentEmployee = Government_Employee.getText();
	   
	WebElement GovtEmp = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_Reservation_AreYouAGovermentEmployee' and @ value='True']")));
	act.moveToElement(GovtEmp).click().perform();

    WebElement dateofjoining = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Date of Joining')]")));
    String dateof_joining = dateofjoining.getText();
	    
    WebElement GovernmentDepartment = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Government Department')]")));
    String Government_Department = GovernmentDepartment.getText();
    jss.executeScript("arguments[0].scrollIntoView(true);", GovernmentDepartment);

    WebElement Designation = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Designation in Government Department')]")));
    String Designationgovt = Designation.getText();
    
    WebElement DepartmentalEnquiry = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Have you been involved in any Departmental Enquiry')]")));
    String Departmental_Enquiry = DepartmentalEnquiry.getText(); 
    

	WebElement DeptEnquiry = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_CriminalActivity_HasDepartmentEnquiry' and @ value='True']")));
	act.moveToElement(DeptEnquiry).click().perform();

    WebElement details = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'If Yes, mention the details')]")));
    String detail = details.getText();
    
    WebElement CriminalCases = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Have you been involved in any Criminal Cases')]")));
    String Criminal_Cases = CriminalCases.getText();
    
    WebElement convicted = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Have you been convicted in a Criminal Case?')]")));
    String convicte = convicted.getText();
    jss.executeScript("arguments[0].scrollIntoView(true);", convicted);
    //SSLC
    WebElement pass = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Passed in SSLC')]")));
    String pass_sslc = pass.getText();
    
	WebElement SSLC = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_IsSSLCHolder' and @ value='True']")));
	act.moveToElement(SSLC).click().perform();
	
    WebElement Boards = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'SSLC Board')]")));
    String board = Boards.getText();
    
    WebElement Board  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_QualificationBoardCode")));
	Select s17=new Select(Board );
	s17.selectByIndex(3);
    
    WebElement other = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'If Selected other board, Enter the Board Name')]")));
    String otherboard = other.getText();

    WebElement Kannada_Language = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'In SSLC, have you studied Kannada Language as')]")));
    String KannadaLanguage = Kannada_Language.getText();
    jss.executeScript("arguments[0].scrollIntoView(true);", Kannada_Language);

    WebElement yearofpass = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Year of Passing SSLC')]")));
    String yearpass = yearofpass.getText();

    WebElement markorgrade = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Marks or Grade')]")));
    String MarkorGrade = markorgrade.getText();
  
    WebElement Markorgrade = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_SSLCQualification_MarkType' and @ value='M']")));
	act.moveToElement(Markorgrade).click().perform();
	
	WebElement Max = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Maximum marks in SSLC')]")));
    String maximum = Max.getText();

    WebElement Obtaine = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Marks Obtained in SSLC')]")));
    String Obtained = Obtaine.getText();

    WebElement Percentage = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Percentage/CGPA in SSLC')]")));
    String percentage = Percentage.getText();

    WebElement Markorgrades = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_SSLCQualification_MarkType' and @ value='G']")));
   	act.moveToElement(Markorgrades).click().perform();
    
   	WebElement grade = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Grade Obtained in SSLC')]")));
    String grades = grade.getText();

    WebElement Register = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'SSLC Registration Number')]")));
    String Registeration = Register.getText();
    jss.executeScript("arguments[0].scrollIntoView(true);", Register);
    
    //puc
    
    WebElement pu = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Do you possess PUC Qualification')]")));
    String puc = pu.getText();

	WebElement Puc = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_IsPUCHolder' and @ value='True']")));
	act.moveToElement(Puc).click().perform();	

    WebElement puboard = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'PUC Board')]")));
    String pucboard = puboard.getText();
    jss.executeScript("arguments[0].scrollIntoView(true);", puboard);
    WebElement PUBoard  = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_QualificationBoardCode")));
	Select s20=new Select(PUBoard );
	s20.selectByIndex(3);
	
    WebElement otherpu = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'If Selected other board, Enter the Board Name')]")));
    String otherpuc = otherpu.getText();
	    
    WebElement puyear = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Year of Passing PUC')]")));
    String pucyear = puyear.getText();
    

    WebElement mg = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Marks or Grade')]")));
    String MG = mg.getText();
    
    WebElement PUMarkorgrade = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_PUCQualification_MarkType' and @ value='M']")));
	act.moveToElement(PUMarkorgrade).click().perform();

    WebElement pumax = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Maximum marks in PUC')]")));
    String pucmax = pumax.getText();

    WebElement puob = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Marks Obtained in PUC')]")));
    String pucob = puob.getText();

    WebElement per = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Percentage / CGPA in PUC')]")));
    String Per = per.getText();
    
    WebElement PUMarkorgrades = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_PUCQualification_MarkType' and @ value='G']")));
	act.moveToElement(PUMarkorgrades).click().perform();

    WebElement pugrade = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Grade Obtained in PUC')]")));
    String puGrade = pugrade.getText();
    jss.executeScript("arguments[0].scrollIntoView(true);", pugrade);
    
    WebElement pureg = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'PUC Registration Number')]")));
    String pucreg = pureg.getText();
    
   //degree
    WebElement degree = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Do you possess degree Qualification')]")));
    String Degree = degree.getText();
    jss.executeScript("arguments[0].scrollIntoView(true);", degree);
    
    //document
    WebElement id = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Select the Uploading Identity Card')]")));
    String idcard = id.getText();
    
    WebElement idno = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Identity Card No')]")));
    String idcardno = idno.getText();
    
    WebElement mark = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Identification mark-01')]")));
    String mark1 = mark.getText();
    
    WebElement marks = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Identification mark-02')]")));
    String marks1 = marks.getText();
    jss.executeScript("arguments[0].scrollIntoView(true);", marks);
    
    WebElement photo = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h1[contains(text(),'Applicant Photo')]")));
    String photos = photo.getText();
    
    WebElement sign = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h1[contains(text(),'Applicant Signature')]")));
    String signature = sign.getText();
    
    WebElement thumb = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h1[contains(text(),'Applicant Left Thumb Impression')]")));
    String thumbh = thumb.getText();
    
    WebElement appid = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h1[contains(text(),'Applicant ID Card')]")));
    String appidcard = appid.getText();
    
    
    FileInputStream fis1 = new FileInputStream("D://APC//APC_NKK.xlsx");//"D:\steno\TestDataAPC.xlsx"
	XSSFWorkbook workbook = new XSSFWorkbook(fis1);
	 Sheet sheet = workbook.getSheetAt(1);
	 
	 int rownumber = 2;
	 Row row1 = sheet.createRow(rownumber);
	 row1.createCell(1).setCellValue(Applying_as_a);
	 row1.createCell(2).setCellValue(unit_name);
	 row1.createCell(3).setCellValue(Applicant_Full_name);
	 row1.createCell(4).setCellValue(Father_name);
	 row1.createCell(5).setCellValue(Mother_name);
	 row1.createCell(6).setCellValue(Email_id);
	 row1.createCell(7).setCellValue(Mobile_no);
	 row1.createCell(8).setCellValue(Aadhar_no);
	 row1.createCell(9).setCellValue(dob);
	 row1.createCell(10).setCellValue(age_applicant);
	 row1.createCell(11).setCellValue(gen);
	 row1.createCell(12).setCellValue(height_applicant);
	 row1.createCell(13).setCellValue(door_no);
	 row1.createCell(14).setCellValue(street);
	 row1.createCell(15).setCellValue(taluk);
	 row1.createCell(16).setCellValue(city);
	 row1.createCell(17).setCellValue(state);
	 row1.createCell(18).setCellValue(district);
	 row1.createCell(19).setCellValue(other_district);
	 row1.createCell(20).setCellValue(Pincode);
	 row1.createCell(21).setCellValue(landmark);
	 row1.createCell(22).setCellValue(Native_dis);
	 row1.createCell(23).setCellValue(peradd);
	 row1.createCell(24).setCellValue(door_no);
	 row1.createCell(25).setCellValue(street);
	 row1.createCell(26).setCellValue(taluk);
	 row1.createCell(27).setCellValue(city);
	 row1.createCell(28).setCellValue(state);
	 row1.createCell(29).setCellValue(district);
	 row1.createCell(30).setCellValue(other_district);
	 row1.createCell(31).setCellValue(Pincode);
	 row1.createCell(32).setCellValue(landmark);
	 row1.createCell(33).setCellValue(category);
	 row1.createCell(34).setCellValue(SubCaste);
	 row1.createCell(35).setCellValue(Caste_Certificate);
	 row1.createCell(36).setCellValue(Kannada_Medium);
	 row1.createCell(37).setCellValue(pdp);
	 row1.createCell(38).setCellValue(belong_to_a_family);
	 row1.createCell(39).setCellValue(Deathwhile_in_service);
	 row1.createCell(40).setCellValue(Permanentlydisabled);
	 row1.createCell(41).setCellValue(Relationship_Exservice);
	 row1.createCell(42).setCellValue(Ex_Servicemen);
	 row1.createCell(43).setCellValue(Armed_Force);
	 row1.createCell(44).setCellValue(benefitexs);
	 row1.createCell(45).setCellValue(Dischargedate);
	 row1.createCell(46).setCellValue(Educational_Qualification);
	 row1.createCell(47).setCellValue(Servicerendered);
	 row1.createCell(48).setCellValue(Serviceyear);
	 row1.createCell(49).setCellValue(years);
	 row1.createCell(50).setCellValue(Month);
	 row1.createCell(51).setCellValue(Day);
	 row1.createCell(52).setCellValue(Transgender_ser);
	 row1.createCell(53).setCellValue(Rural_medium);
	
	 row1.createCell(56).setCellValue(GovernmentEmployee);
	 row1.createCell(57).setCellValue(dateof_joining);
	 row1.createCell(58).setCellValue(Serviceyear);
	 row1.createCell(59).setCellValue(years);
	 row1.createCell(60).setCellValue(Month);
	 row1.createCell(61).setCellValue(Day);
	 row1.createCell(62).setCellValue(Government_Department);
	 row1.createCell(63).setCellValue(Designationgovt);
	 row1.createCell(64).setCellValue(Departmental_Enquiry);
	 row1.createCell(65).setCellValue(detail);
	 row1.createCell(66).setCellValue(Criminal_Cases);
	 row1.createCell(67).setCellValue(detail);
	 row1.createCell(68).setCellValue(convicte);
	 row1.createCell(69).setCellValue(detail);
	 row1.createCell(70).setCellValue(pass_sslc);
	 row1.createCell(71).setCellValue(board);
	 row1.createCell(72).setCellValue(otherboard);
	 row1.createCell(73).setCellValue(KannadaLanguage);
	 row1.createCell(74).setCellValue(yearpass);
	 row1.createCell(75).setCellValue(MarkorGrade);
	 row1.createCell(76).setCellValue(maximum);
	 row1.createCell(77).setCellValue(Obtained);
	 row1.createCell(78).setCellValue(percentage);
	 row1.createCell(79).setCellValue(grades);
	 row1.createCell(80).setCellValue(percentage);
	 row1.createCell(81).setCellValue(Registeration);
     row1.createCell(82).setCellValue(puc);
	 row1.createCell(83).setCellValue(pucboard);
	 row1.createCell(84).setCellValue(otherpuc);
	 row1.createCell(85).setCellValue(pucyear);
	 row1.createCell(86).setCellValue(MG);
	 row1.createCell(87).setCellValue(pucmax);
	 row1.createCell(88).setCellValue(pucob);
	 row1.createCell(89).setCellValue(Per);
	 row1.createCell(90).setCellValue(puGrade);
	 row1.createCell(91).setCellValue(Per);
	 row1.createCell(92).setCellValue(pucreg);
	 row1.createCell(93).setCellValue(Degree);
	 row1.createCell(94).setCellValue(idcard);
	 row1.createCell(95).setCellValue(idcardno);
	 row1.createCell(96).setCellValue(mark1);
	 row1.createCell(97).setCellValue(marks1);
	 row1.createCell(98).setCellValue(photos);
	 row1.createCell(99).setCellValue(signature);
	 row1.createCell(100).setCellValue(thumbh);
	 row1.createCell(101).setCellValue(appidcard);


  FileOutputStream file = new FileOutputStream("D://Automation_data//TestDataAPC_NKK.xlsx");
  workbook.write(file);
  file.close();
  driver.quit();  
}
private void switchToNewWindow(WebDriver driver) {
	Set<String> window = driver.getWindowHandles();
	for (String handle : window) {
		driver.switchTo().window(handle);
	}
	
}
}
