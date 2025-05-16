package Steno;

import org.testng.Reporter;
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
import org.testng.annotations.Test;

public class New_app_title {
	
    @Test(priority=1)
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
        Thread.sleep(1000);

        js.executeScript("window.scrollBy(0,200);");
        Thread.sleep(1000);
       

		try
		{
		WebElement dpd = driver.findElement(By.id("Applicant_PostUnitCode"));
		// dpd.click();
		Select ss = new Select(dpd);
		Thread.sleep(2000);
		ss.selectByIndex(1); 
		
       WebElement app = driver.findElement(By.xpath("//label[contains(text(), ' Applying Post ')]"));
       String post = app.getText();
    
	   js.executeScript("arguments[0].scrollIntoView(true);", app);

       //Applicant details
       WebElement appn = driver.findElement(By.xpath("//div[contains(text(),'Applicant Full Name ')]"));
        String name = appn.getText();
  
        
        WebElement father = driver.findElement(By.xpath("//div[contains(text(),'Father')]"));
        String fathername = father.getText();
   
         
        WebElement mother = driver.findElement(By.xpath("//div[contains(text(),'Mother')]"));
        String mothername = mother.getText();
    
 	   js.executeScript("arguments[0].scrollIntoView(true);", mother);

        
        WebElement email = driver.findElement(By.xpath("//div[contains(text(),'Email ID')]"));
        String emailid = email.getText();
        
        WebElement mobile = driver.findElement(By.xpath("//div[contains(text(),'Mobile')]"));
        String mobileno = mobile.getText();
     
  	   js.executeScript("arguments[0].scrollIntoView(true);", mobile);
  	 Actions act=new Actions(driver);
        
        WebElement adhar = driver.findElement(By.xpath("//div[contains(text(),'Aadhar ')]"));
        String adharno = adhar.getText();
     
        WebElement dob = driver.findElement(By.xpath("//div[contains(text(),'Date of Birth')]"));
        String Dob = dob.getText();
     
   	   js.executeScript("arguments[0].scrollIntoView(true);", dob);

        WebElement age = driver.findElement(By.xpath("//div[contains(text(),'Age of Applicant as on 04th July 2024 ')]"));
        String ageason = age.getText();
      
        WebElement gender = driver.findElement(By.xpath("//div[contains(text(),'Gender ')]"));
        String gen = gender.getText();
       
        //address
        WebElement door = driver.findElement(By.xpath("//div[contains(text(),'Door ')]"));
        String doorno = door.getText();
      	js.executeScript("arguments[0].scrollIntoView(true);", door);

        WebElement stret = driver.findElement(By.xpath("//div[contains(text(),'Street')]"));
        String street = stret.getText();
      
        
        WebElement land = driver.findElement(By.xpath("//div[contains(text(),'Nearby Landmark')]"));
        String landmark = land.getText();
      
    	//js.executeScript("arguments[0].scrollIntoView(true);", land);

        WebElement tlk = driver.findElement(By.xpath("//div[contains(text(),'Taluk ')]"));
        String taluk = tlk.getText();

        
        WebElement citi = driver.findElement(By.xpath("//div[contains(text(),'City ')]"));
        String city = citi.getText();
      

        WebElement stat = driver.findElement(By.xpath("//div[contains(text(),'State ')]"));
        String state = stat.getText();
    
        WebElement sta = driver.findElement(By.id("Applicant_ContactAddress_UnionStateCode"));
		Thread.sleep(500);
		act.moveToElement(sta).click().perform();
		Thread.sleep(500);
		Select ss30 = new Select(sta);
		Thread.sleep(500);
		ss30.selectByIndex(14);

		WebElement dis = driver.findElement(By.id("Applicant_ContactAddress_DistrictCode"));
		Thread.sleep(500);
		act.moveToElement(dis).click().perform();
		Thread.sleep(500);
		Select ss3 = new Select(dis);
		Thread.sleep(500);
		ss3.selectByIndex(32);
        WebElement dist = driver.findElement(By.xpath("//div[contains(text(),'District ')]"));
        String district = dist.getText();
   
        
        
        WebElement oth = driver.findElement(By.xpath("//div[contains(text(),'If selected other district mention the Specific ')]"));
        String other = oth.getText();
       
        WebElement pin = driver.findElement(By.xpath("//div[contains(text(),'Pincode ')]"));
        String pincode = pin.getText();
     	js.executeScript("arguments[0].scrollIntoView(true);", pin);

        
        WebElement ndist = driver.findElement(By.xpath("//div[contains(text(),'Native District')]"));
        String nativedist = ndist.getText();
   
        WebElement add = driver.findElement(By.xpath("//div[contains(text(),'Postal Address and Permanent Address are Same')]"));
        String addssame = add.getText();
        
        
        //Applicant_Reservation details
        WebElement cat = driver.findElement(By.xpath("//div[contains(text(),'Category ')]"));
        String catogary = cat.getText();
          
        WebElement cast = driver.findElement(By.id("Applicant_Reservation_CategoryCode"));
        Thread.sleep(1000);
		act.moveToElement(cast).click().perform();
		Thread.sleep(1000);
		Select ss4 = new Select(cast);
		ss4.selectByIndex(1);
        
        WebElement sub = driver.findElement(By.xpath("//div[contains(text(),'Sub Caste')]"));
        String subcaste = sub.getText();
    
        Thread.sleep(500);
        WebElement dateofissue = driver.findElement(By.xpath("//div[contains(text(),'Date of Issue of caste certificate')]"));
        String dateofissuesub = dateofissue.getText();
      
    	js.executeScript("arguments[0].scrollIntoView(true);", dateofissue);

        WebElement emp = driver.findElement(By.xpath("//div[contains(text(),'Are you a Government Employee?')]"));
        String govtemp = emp.getText();
         
        driver.findElement(By.xpath("//input[@id='Applicant_Reservation_AreYouAGovermentEmployee' and @value='True']")).click();
		Thread.sleep(500);
        WebElement doj = driver.findElement(By.xpath("//div[contains(text(),'Date of Joining?')]"));
        String dojoin = doj.getText();
          
        WebElement govtdep = driver.findElement(By.xpath("//div[contains(text(),'Government Department :')]"));
        String govtdept = govtdep.getText();
        
        WebElement ser = driver.findElement(By.xpath("//div[contains(text(),'How many years of Service have you rendered?')]"));
        String service = ser.getText();
     
        js.executeScript("arguments[0].scrollIntoView(true);", ser);
        
        WebElement year = driver.findElement(By.xpath("//label[contains(text(),'Years /')]"));
        String years = year.getText();
   
        
        WebElement moth = driver.findElement(By.xpath("//label[contains(text(),'Months /')]"));
        String moths = moth.getText();

        
        WebElement day = driver.findElement(By.xpath("//div[contains(text(),'Government Department :')]"));
        String days = day.getText();
     
        
        WebElement disg = driver.findElement(By.xpath("//div[contains(text(),'Designation in Government Department?')]"));
        String disgna = disg.getText();
    
        
        WebElement depen = driver.findElement(By.xpath("//div[contains(text(),'Have you been involved in any Departmental Enquiry?')]"));
        String deptenq =depen.getText();
      
        //js.executeScript("arguments[0].scrollIntoView(true);", depen);
        Thread.sleep(500);
        
        WebElement ab = driver.findElement(By.xpath("//input[@id='Applicant_CriminalActivity_HasDepartmentEnquiry' and @value='True']"));
        Thread.sleep(500);
		act.moveToElement(ab).click().perform();
		
		
		 WebElement depend = driver.findElement(By.xpath("//div[contains(text(),'If Yes, mention the details :')]"));
	     String deptenqde =depend.getText();
   

	     WebElement cr = driver.findElement(By.xpath("//input[@id='Applicant_CriminalActivity_IsInvolvedInCriminalActivity' and @value='True']"));
		act.moveToElement(cr).click().perform();
		Thread.sleep(500);
		
		  WebElement crm = driver.findElement(By.xpath("//div[contains(text(),'Have you been involved in any Criminal Cases?')]"));
	      String crime = crm.getText();

	      js.executeScript("arguments[0].scrollIntoView(true);", crm);
	      WebElement depend1 = driver.findElement(By.xpath("//div[contains(text(),'If Yes, mention the details :')]"));
		  String deptenqde1 =depend1.getText();
		
		  
		  WebElement crd = driver.findElement(By.xpath("//input[@id='Applicant_CriminalActivity_IsConvictedInCriminalCase' and @value='True']"));
		  act.moveToElement(crd).click().perform(); // Move to the element and click it
			Thread.sleep(2000);
			
			 WebElement con = driver.findElement(By.xpath("//div[contains(text(),'Have you been convicted in a Criminal Case?')]"));
		      String conv = con.getText();
		 
		      
		      WebElement depend2 = driver.findElement(By.xpath("//div[contains(text(),'If Yes, mention the details :')]"));
			  String deptenqde2 =depend2.getText();
				  
			  // for sslc
			  WebElement sslcc = driver.findElement(By.xpath("//div[contains(text(),'Passed in SSLC / Equivalent?')]"));
		        String sslcp = sslcc.getText();
		        
		        WebElement sslc = driver.findElement(By.xpath("//input[@id='Applicant_EducationalQualification_IsSSLCHolder' and @value='True']"));
				Thread.sleep(1000);
				act.moveToElement(sslc).click().perform();
				
				
		        WebElement boarde = driver.findElement(By.xpath("//div[contains(text(),'SSLC Board :')]"));
		        String boards = boarde.getText();
		        
		        WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(20));
		    	WebElement board = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_SSLCQualification_QualificationBoardCode")));
		    	wait1.until(ExpectedConditions.visibilityOf(board));
				wait1.until(ExpectedConditions.elementToBeClickable(board));
				Select ss8 = new Select(board);
				ss8.selectByIndex(3);
				 
		        WebElement sslcoth = driver.findElement(By.xpath("//div[contains(text(),'If Selected other board, Enter the Board Name ')]"));
		        String sslcother = sslcoth.getText();
				Thread.sleep(200);
		        js.executeScript("window.scrollBy(0,100);");
		        
		        WebElement kannada = driver.findElement(By.xpath("//div[contains(text(),'n SSLC, have you studied Kannada Language ')]"));
		        String kannadalang = kannada.getText();
		        //js.executeScript("arguments[0].scrollIntoView(true);", kannada);
		        
		        WebElement yearpass = driver.findElement(By.xpath("//div[contains(text(),'Year of Passing SSLC:')]"));
		        String yearofpass = yearpass.getText();
		       
		        
		        
		        WebElement mg = driver.findElement(By.xpath("//div[contains(text(),'Marks or Grade?')]"));
		        String markorgrade = mg.getText();
		        js.executeScript("window.scrollBy(0,200);");
		        
		    	WebElement mark = wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_SSLCQualification_MarkType' and @value='M']")));
				Thread.sleep(1000);
				act.moveToElement(mark).click().perform();
			   
		        WebElement maxmark = driver.findElement(By.xpath("//div[contains(text(),'Maximum marks in SSLC :')]"));
		        String maxmarks = maxmark.getText();
				Thread.sleep(200);
		     //   js.executeScript("arguments[0].scrollIntoView(true);", maxmark);
		        
		        WebElement minmark = driver.findElement(By.xpath("//div[contains(text(),'Marks Obtained in SSLC :')]"));
		        String minmarks = minmark.getText();
				Thread.sleep(200);
		     //   js.executeScript("arguments[0].scrollIntoView(true);", minmark);
		        
		        WebElement gradee = wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_SSLCQualification_MarkType' and @value='G']")));
				Thread.sleep(500);
				act.moveToElement(gradee).click().perform();
		        WebElement grade = driver.findElement(By.xpath("//div[contains(text(),'Grade Obtained in SSLC :')]"));
		        String gradeob = grade.getText();
		        js.executeScript("window.scrollBy(0,200);");
		        
		        WebElement per = driver.findElement(By.xpath("//div[contains(text(),'Percentage/CGPA in SSLC :')]"));
		        String percen = per.getText();
				Thread.sleep(500);
		     //   js.executeScript("arguments[0].scrollIntoView(true);", per);
		        
		        WebElement reg = driver.findElement(By.xpath("//div[contains(text(),'SSLC Registration Number :')]"));
		        String regno = reg.getText();
		    
				Thread.sleep(500);
		       // js.executeScript("arguments[0].scrollIntoView(true);", regno);
		        
				// for puc
				
		        WebElement pu = driver.findElement(By.xpath("//div[contains(text(),'Do you possess PUC Qualification?')]"));
		        String puc = pu.getText();
		      
				Thread.sleep(500);
				 js.executeScript("window.scrollBy(0,200);");
			        
		        WebElement pucp = driver.findElement(By.xpath("//input[@id='Applicant_EducationalQualification_IsPUCHolder' and @value='True']"));
				Thread.sleep(500);
				act.moveToElement(pucp).click().perform();
			
				
		        WebElement puboard = driver.findElement(By.xpath("//div[contains(text(),'PUC Board :')]"));
		        String pucboard = puboard.getText();
				Thread.sleep(500);

				WebElement puboard1 = wait1.until(ExpectedConditions.presenceOfElementLocated(By.id("Applicant_EducationalQualification_PUCQualification_QualificationBoardCode")));
				wait1.until(ExpectedConditions.visibilityOf(puboard1));
				wait1.until(ExpectedConditions.elementToBeClickable(puboard1));
				Thread.sleep(500);
				Select ss14 = new Select(puboard1);
				ss14.selectByIndex(3);
				
				WebElement puboardoth = driver.findElement(By.xpath("//div[contains(text(),'If Selected other board, Enter the Board Name ')]"));
			    String pucboardoyh = puboardoth.getText();
			    Thread.sleep(500);

		        WebElement puyear = driver.findElement(By.xpath("//div[contains(text(),'Year of Passing PUC :')]"));
		        String pucyear = puyear.getText();
				Thread.sleep(500);
				 js.executeScript("window.scrollBy(0,200);");
			        
		        WebElement morg = driver.findElement(By.xpath("//div[contains(text(),'Marks or Grade?')]"));
		        String markorg = morg.getText();
				Thread.sleep(500);
				 js.executeScript("window.scrollBy(0,200);");
		        
		        WebElement gradem = wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_PUCQualification_MarkType' and @value='M']")));
				Thread.sleep(500);
				act.moveToElement(gradem).click().perform();
				
		        
		        WebElement pumax = driver.findElement(By.xpath("//div[contains(text(),'Maximum marks in PUC :')]"));
		        String pucmax = pumax.getText();
		      
		        
		        WebElement puobt = driver.findElement(By.xpath("//div[contains(text(),'Marks Obtained in PUC :')]"));
		        String pucobt = puobt.getText();
				Thread.sleep(500);
		        
		        WebElement markpu = wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='Applicant_EducationalQualification_PUCQualification_MarkType' and @value='G']")));
				Thread.sleep(500);
				act.moveToElement(markpu).click().perform();
		        
		        WebElement pucgpa = driver.findElement(By.xpath("//div[contains(text(),'Grade Obtained in PUC :')]"));
		        String puccgpa = pucgpa.getText();
				Thread.sleep(500);
				 js.executeScript("window.scrollBy(0,200);");
			        
		        WebElement puper = driver.findElement(By.xpath("//div[contains(text(),'Percentage / CGPA in PUC :')]"));
		        String pucper = puper.getText();
				Thread.sleep(500);
		        
		        WebElement pureg = driver.findElement(By.xpath("//div[contains(text(),'PUC Registration Number :')]"));
		        String pucreg = pureg.getText();
		        
		        // for degree
		        WebElement deg = driver.findElement(By.xpath("//div[contains(text(),'Are you a degree holder?')]"));
		        String degree = deg.getText();
		        js.executeScript("window.scrollBy(0,200);");
		        
		        //for qualification
		        WebElement typ = driver.findElement(By.xpath("//div[contains(text(),'Have you Passed Senior Kannada Typist Exam?')]"));
		        String typist = typ.getText();
		        
				WebElement senior = driver.findElement(By.xpath("//input[@id='Applicant_TypistAssistant_IsPassedInQualifyExam' and @value='True']"));
				Thread.sleep(500);
				act.moveToElement(senior).click().perform();
				
		        
		        WebElement typopt = driver.findElement(By.xpath("//div[contains(text(),'Select Qualified Examination from the List ')]"));
		        String typop = typopt.getText();
		        
		        WebElement sen = driver.findElement(By.xpath("//div[contains(text(),'Have you Passed Senior Kannada Shorthand Exam?')]"));
		        String senn = sen.getText();
		        
		        WebElement se = driver.findElement(By.xpath("//input[@id='Applicant_StenographerAssistant_IsPassedInQualifyExam' and @value='True']"));
				Thread.sleep(500);
				act.moveToElement(se).click().perform();
			
				js.executeScript("window.scrollBy(0,200);");
		        
		        
		        WebElement opt = driver.findElement(By.xpath("//div[contains(text(),'Select Qualified Examination from the List ')]"));
		        String option = opt.getText();
		        Thread.sleep(500);
		        
		        //for document upload
		        WebElement id = driver.findElement(By.xpath("//div[contains(text(),'Select the Uploading Identity Card :')]"));
		        String idup = id.getText();
		      
		        js.executeScript("window.scrollBy(0,200);");
		        
		        WebElement idno = driver.findElement(By.xpath("//div[contains(text(),'Identity Card Number :')]"));
		        String idnoup = idno.getText();
		  
		      
		        
		        WebElement idm = driver.findElement(By.xpath("//div[contains(text(),'Identification Mark-01 :')]"));
		        String idmk = idm.getText();
		        
		        Thread.sleep(300);
		        js.executeScript("window.scrollBy(0,200);");
		        
		       
		        WebElement idk1 = driver.findElement(By.xpath("//div[contains(text(),'Identification Mark-02 :')]"));
		        String idmk1 = idk1.getText();
		        Thread.sleep(500);
		        
		        WebElement photo = driver.findElement(By.xpath("//h1[contains(text(),'Applicant Photo & Signature')]"));
		        String photoo = photo.getText();
		     
		        Thread.sleep(500);
		        
		        WebElement thumb = driver.findElement(By.xpath("//h1[contains(text(),'Applicant Left Thumb Impression ')]"));
		        String thumbb = thumb.getText();
		    
		      
		        
		        WebElement idp = driver.findElement(By.xpath("//h1[contains(text(),'Applicant ID Card ')]"));
		        String idphoto = id.getText();
		       
		        FileInputStream fis1 = new FileInputStream("D://steno//TestData (2).xlsx");// C://Users//pallavi//eclipse-workspace//project//Book5.xlsx
				XSSFWorkbook workbook = new XSSFWorkbook(fis1);
		   	 Sheet sheet = workbook.getSheetAt(2);
			 //  Row row1 = sheet.createRow(sheet.getPhysicalNumberOfRows()); 
			     int rowNumber = 2; 
				 Row row1 = sheet.createRow(rowNumber); 
				        row1.createCell(1).setCellValue(post);
	                    row1.createCell(2).setCellValue(name);
	                    row1.createCell(3).setCellValue(fathername);
	                    row1.createCell(4).setCellValue(mothername);
	                    row1.createCell(5).setCellValue(emailid);
	                    row1.createCell(6).setCellValue(mobileno);
	                    row1.createCell(7).setCellValue(adharno);
	                    row1.createCell(8).setCellValue(Dob);
	                    row1.createCell(9).setCellValue(ageason);
	                    row1.createCell(10).setCellValue(gen);
	                    row1.createCell(11).setCellValue(doorno);
	                    row1.createCell(12).setCellValue(street);
	                 
	                    row1.createCell(13).setCellValue(taluk);
	                    row1.createCell(14).setCellValue(city);
	                    row1.createCell(15).setCellValue(state);
	                    row1.createCell(16).setCellValue(district);
	                    row1.createCell(17).setCellValue(other);
	                    row1.createCell(18).setCellValue(pincode);
	                    row1.createCell(19).setCellValue(landmark);
	                    row1.createCell(20).setCellValue(nativedist);
	                    row1.createCell(21).setCellValue(addssame);
	                    row1.createCell(22).setCellValue(doorno);
	                    row1.createCell(23).setCellValue(street);
	                    row1.createCell(24).setCellValue(taluk);
	                    row1.createCell(25).setCellValue(city);
	                    row1.createCell(26).setCellValue(state);
	                    row1.createCell(27).setCellValue(district);
	                    row1.createCell(27).setCellValue(other);
	                    row1.createCell(29).setCellValue(pincode);
	                    row1.createCell(30).setCellValue(landmark);
	                    row1.createCell(31).setCellValue(catogary);
	                    row1.createCell(32).setCellValue(subcaste);
	                    row1.createCell(33).setCellValue(dateofissuesub);
	                    row1.createCell(34).setCellValue(govtemp);
	                    row1.createCell(35).setCellValue(dojoin);
	                    row1.createCell(36).setCellValue(govtdept);
	                    row1.createCell(37).setCellValue(service);
	                    row1.createCell(38).setCellValue(years);
	                    row1.createCell(39).setCellValue(moths);
	                    row1.createCell(40).setCellValue(days);
	                    row1.createCell(41).setCellValue(disgna);
	                    row1.createCell(42).setCellValue(deptenq);
	                    row1.createCell(43).setCellValue(deptenqde);
	                    row1.createCell(44).setCellValue(crime);
	                    row1.createCell(45).setCellValue(deptenqde1);
	                    row1.createCell(46).setCellValue(conv);
	                    row1.createCell(47).setCellValue(deptenqde1);
	                    row1.createCell(48).setCellValue(sslcp);
	                    row1.createCell(49).setCellValue(boards);
	                    row1.createCell(50).setCellValue(sslcother);
	                    row1.createCell(51).setCellValue(kannadalang);
	                    row1.createCell(52).setCellValue(yearofpass);
	                    row1.createCell(53).setCellValue(markorgrade);
	                    row1.createCell(54).setCellValue(maxmarks);
	                    row1.createCell(55).setCellValue(minmarks);
	                    row1.createCell(56).setCellValue(percen);
	                    row1.createCell(57).setCellValue(gradeob);
	                    row1.createCell(58).setCellValue(percen);
	                    
	                    
	                    row1.createCell(59).setCellValue(regno);
	                    row1.createCell(60).setCellValue(puc);
	                    row1.createCell(61).setCellValue(pucboard);
	                    row1.createCell(62).setCellValue(pucboardoyh);
	                    row1.createCell(63).setCellValue(pucyear);
	                    row1.createCell(64).setCellValue(markorg);
	                    row1.createCell(65).setCellValue(pucmax);
	                    row1.createCell(66).setCellValue(pucobt);
	                    row1.createCell(67).setCellValue(pucper);
	                    row1.createCell(68).setCellValue(puccgpa);
	                    row1.createCell(69).setCellValue(pucper);
	                    row1.createCell(70).setCellValue(pucreg);
	                    row1.createCell(71).setCellValue(degree);
	                    row1.createCell(72).setCellValue(typist);
	                    row1.createCell(73).setCellValue(typop);
	                    row1.createCell(74).setCellValue(senn); 
	                    row1.createCell(75).setCellValue(option);
	                   
	                    row1.createCell(76).setCellValue(idup);
	                    row1.createCell(77).setCellValue(idnoup);
	                    row1.createCell(78).setCellValue(idmk);
	                    row1.createCell(79).setCellValue(idmk1);
	                    row1.createCell(80).setCellValue(photoo);
	                    row1.createCell(81).setCellValue(thumbb);
	                    row1.createCell(82).setCellValue(idphoto);

	                    FileOutputStream fileOut = new FileOutputStream("D://steno//TestData (2).xlsx");
	    	            workbook.write(fileOut);
	    	            fileOut.close();
					    Reporter.log("New Application Titles are completed");
	    	           
	    	            

	    	            if (driver != null) {
	    	                driver.quit();
	    	                
	    	            }}
	    			 catch (IOException e) {
	    	            e.printStackTrace();
	    	        } }
				
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
				@Test(priority=2)
				public void sample2() throws InterruptedException, AWTException, IOException
				{
					ChromeDriver driver = new ChromeDriver();
					driver.get("http://172.10.1.159:9013/Auth/login");
					driver.manage().window().maximize();
				    WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(20));//
				    WebElement app_no = driver.findElement(By.xpath("//label[contains(text(),'Application Number')]"));
				    String ap_no = app_no.getText();
				    
				    WebElement do_no = driver.findElement(By.xpath("//label[contains(text(),'Date of Birth')]"));
				    String dob = do_no.getText();
				    
				    WebElement for_app = driver.findElement(By.xpath("//a[contains(text(),'Forgot Application Number?')]"));
				    String forgot = for_app.getText();
				    
				    WebElement log = driver.findElement(By.xpath("//button[contains(text(),'Login')]"));
				    String login = log.getText();
				    
				    driver.findElement(By.id("ApplicantModel_ApplicationNo")).sendKeys("0000743");
				    Thread.sleep(1000);
				    driver.findElement(By.id("ApplicantModel_DateOfBirth")).sendKeys("10-12-1999");
				    
				    
					Robot rb=new Robot();
					rb.keyPress(KeyEvent.VK_ENTER);
					rb.keyRelease(KeyEvent.VK_ENTER);
					
					Thread.sleep(1000);
					driver.findElement(By.id("login_submit")).click();
					Thread.sleep(4000);
					Set<String> window = driver.getWindowHandles();
					for (String all : window) {
						driver.switchTo().window(all);
					}
					Thread.sleep(1000);
					 JavascriptExecutor jss = (JavascriptExecutor) driver;
					jss.executeScript("window.scrollBy(0,400)");
					
					String applicantName = driver.findElement(By.xpath("//p[contains(text(),'Applicant Name')]")).getText();
					

					String applicationNumber = driver.findElement(By.xpath("//p[contains(text(),'Application Number')]")).getText();

					String dobc = driver.findElement(By.xpath("//p[contains(text(),'Date of Birth ')]")).getText();

					String gender = driver.findElement(By.xpath("//p[contains(text(),'Gender')]")).getText();

					String aadharNumber = driver.findElement(By.xpath("//p[contains(text(),'Aadhar Number')]")).getText();


					String mobileNumber = driver.findElement(By.xpath("//p[contains(text(),'Mobile Number')]")).getText();

					String category = driver.findElement(By.xpath("//p[contains(text(),'Category')]")).getText();


					String email = driver.findElement(By.xpath("//p[contains(text(),' E-mail')]")).getText();
					
					String photo = driver.findElement(By.xpath("//p[contains(text(),'Applicant Photo & Signature')]")).getText();

					String id = driver.findElement(By.xpath("//p[contains(text(),'Applicant ID Card')]")).getText();


					String thumb = driver.findElement(By.xpath("//p[contains(text(),'Thumb')]")).getText();

				    FileInputStream fis1 = new FileInputStream("D://steno//TestData (2).xlsx");// C://Users//pallavi//eclipse-workspace//project//Book5.xlsx
					XSSFWorkbook workbook = new XSSFWorkbook(fis1);
					 Sheet sheet = workbook.getSheetAt(2);
					 int rowNumber = 9; 
					 Row row1 = sheet.createRow(rowNumber);
				  // Row row1 = sheet.createRow(sheet.getPhysicalNumberOfRows()); 
					   
					        row1.createCell(1).setCellValue(ap_no);
				            row1.createCell(2).setCellValue(dob);
				            row1.createCell(3).setCellValue(forgot);
				            row1.createCell(4).setCellValue(login);
				            row1.createCell(5).setCellValue(applicantName);
				            row1.createCell(6).setCellValue(applicationNumber);
				            row1.createCell(7).setCellValue(dobc);
				            row1.createCell(8).setCellValue(gender);
				            row1.createCell(9).setCellValue(aadharNumber);
				            row1.createCell(10).setCellValue(mobileNumber);
				            row1.createCell(11).setCellValue(category);
				            row1.createCell(12).setCellValue(email);
				         
				            row1.createCell(13).setCellValue(photo);
				            row1.createCell(14).setCellValue(id);
				            row1.createCell(15).setCellValue(thumb);
				            
				         
				            
				            
				            FileOutputStream fileOut = new FileOutputStream("D://steno//TestData (2).xlsx");
				            workbook.write(fileOut);
				            fileOut.close();
						    Reporter.log("My Application Titles are completed");

				           driver.quit();
				}
	}


