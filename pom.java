package ksp_admin;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.Duration;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class pom extends Utilies {

    WebDriver driver;
    WebDriverWait wait;
    Robot robot;
    Actions actions;

    
    @FindBy(id = "ApplyingType_Code")
    WebElement codeField;

    @FindBy(id = "ApplyingType_Title")
    WebElement titleField;

    @FindBy(id = "ApplyingType_OrderIndex")
    WebElement orderIndexDropdown;

    @FindBy(id = "ApplyingType_StatusCode")
    WebElement statusDropdown;

    @FindBy(xpath = "//button[contains(text(),'Cancel')]")
    WebElement cancelButton;

    @FindBy(id = "search-bar")
    WebElement searchBar;

    @FindBy(xpath = "//span[contains(text(),'visibility')]")
    WebElement viewButton;

    @FindBy(xpath = "//span[contains(text(),'edit')]")
    WebElement editButton;

    @FindBy(linkText = "arrow_back")
    WebElement backButton;

    @FindBy(id="ApplyingTypeDTO_Code")
    WebElement Editcode;
    
    @FindBy(id="ApplyingTypeDTO_Title")
    WebElement Edittitle;
    
    @FindBy(id="ApplyingTypeDTO_OrderIndex")
    WebElement Editorderindex;
    
    @FindBy(id="ApplyingTypeDTO_StatusCode")
    WebElement Editstatus;
    
    @FindBy(xpath = "//button[contains(text(),'Save')]")
    WebElement save;
    
    @FindBy(xpath = "(//span[contains(text(),'delete')])[1]")
    WebElement delete;
    
    @FindBy(xpath="//button[contains(text(),'OK')]")
    WebElement ok;
    
    @FindBy(xpath="//span[contains(text(),'close')]")
    WebElement clear;
    
    @FindBy(xpath = "(//input[@type='checkbox'])[2]")
    WebElement checkbox1;
    
    @FindBy(xpath = "(//input[@type='checkbox'])[1]")
    WebElement checkboxmulti;
    
    @FindBy(id = "statusSelect")
    WebElement Status;
    
    @FindBy(xpath="//button[contains(text(),'Apply')]")
    WebElement Apply;
    
    @FindBy(xpath="//a[contains(text(),'Add')]")
    WebElement add;
    
    public pom(WebDriver driver) throws AWTException {
        this.driver = driver;
        this.wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        this.robot = new Robot();
        this.actions = new Actions(driver);
      
        PageFactory.initElements(driver, this);
        
    }

    public void openMasters() {
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Masters')]"))).click();
        sleep(1000);
    }

    public void openApplyingTypes() {
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Applying Types')]"))).click();
        sleep(1000);
    }

    public void switchWindow() {
        switchToNewWindow(driver);
        sleep(1000);
    }

    public void clickAdd() {
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Add')]"))).click();
        sleep(1000);
    }

    public void addd() {
    	
    	 FileInputStream fis = null;
 	    FileOutputStream fileOut = null;
        try
              {
            FileInputStream fis1 = new FileInputStream("D://KSP_Admin//ksp_automation.xlsx");//"D:\steno\TestDataAPC.xlsx"
	    	XSSFWorkbook workbook = new XSSFWorkbook(fis1);
	    	Sheet sheet = workbook.getSheetAt(0);
         
           

            for (int i = 1; i <= 2; i++) {
            	
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String code = getCellValue(row.getCell(1));
                String title = getCellValue(row.getCell(2));
                int orderIndex = (int) row.getCell(3).getNumericCellValue();
                int status = (int) row.getCell(4).getNumericCellValue();
                
                String codeEdit = getCellValue(row.getCell(6));
                String titleEdit = getCellValue(row.getCell(7));
                int orderIndexEdit = (int) row.getCell(8).getNumericCellValue();
                int statusEdit = (int) row.getCell(9).getNumericCellValue();

                add.click();
                switchToNewWindow(driver);
                
                codeField.sendKeys(code);
                titleField.sendKeys(title);
                new Select(orderIndexDropdown).selectByIndex(orderIndex);
                new Select(statusDropdown).selectByIndex(status);

                actions.moveToElement(save).click().perform();

                if(isElementClickable(driver, ok)) {
                	try {
                		 wait.until(ExpectedConditions.visibilityOf(ok));
                         ok.click();
                         sleep(1000);
                         cancelButton.click();
                         sleep(500);
                         switchToNewWindow(driver);
                	}
                	catch (Exception r) {
						// TODO: handle exception
					}
                }
                sleep(1000);
                switchToNewWindow(driver);
                actions.moveToElement(searchBar).click().perform();
                searchBar.sendKeys(code);

                robot.keyPress(KeyEvent.VK_ENTER);
                robot.keyRelease(KeyEvent.VK_ENTER);
                
                sleep(1000);
                wait.until(ExpectedConditions.visibilityOf(viewButton));
                actions.moveToElement(viewButton).click().perform();

                switchToNewWindow(driver);
        

                String getViewCode = getAttributeValue(By.id("code"));
                String getViewTitle = getAttributeValue(By.id("title"));
                String getOrderIndex = getAttributeValue(By.id("orderIndex"));
                String getStatus = getAttributeValue(By.id("status"));

               
                actions.moveToElement(backButton).click().perform();
                switchToNewWindow(driver);

                Sheet sheet2 = workbook.getSheetAt(1);
                Row row1 = sheet2.createRow(sheet2.getPhysicalNumberOfRows());
                row1.createCell(1).setCellValue(getViewCode);
                row1.createCell(2).setCellValue(getViewTitle);
                row1.createCell(3).setCellValue(getOrderIndex);
                row1.createCell(4).setCellValue(getStatus);

           
                actions.moveToElement(searchBar).click().perform();
                searchBar.sendKeys(code);
                robot.keyPress(KeyEvent.VK_ENTER);
                robot.keyRelease(KeyEvent.VK_ENTER);
            
                sleep(500);
                actions.moveToElement(editButton).click().perform();
                switchToNewWindow(driver);
              
                
                Editcode.clear();
                Editcode.sendKeys(codeEdit);
                
                Edittitle.clear();
                Edittitle.sendKeys(titleEdit);
                
                Select s = new Select(Editorderindex);
                s.selectByIndex(orderIndexEdit);
                new Select(Editstatus).selectByIndex(statusEdit);
                
                wait.until(ExpectedConditions.visibilityOf(save));
                save.click();
                
                wait.until(ExpectedConditions.visibilityOf(ok));
                ok.click();
              
                
                if(isElementClickable(driver, cancelButton)) {
                	try {
                		
                         sleep(500);
                         cancelButton.click();
                         sleep(500);
                         switchToNewWindow(driver);
                	}
                	catch (Exception r) {
						// TODO: handle exception
                		  wait.until(ExpectedConditions.visibilityOf(ok));
                          ok.click();
					}
                }
                switchToNewWindow(driver);
                sleep(1000);
                wait.until(ExpectedConditions.visibilityOf(searchBar));
                actions.moveToElement(searchBar).click().perform();
                searchBar.sendKeys(codeEdit);
                robot.keyPress(KeyEvent.VK_ENTER);
                robot.keyRelease(KeyEvent.VK_ENTER);
               
                sleep(500);
                wait.until(ExpectedConditions.visibilityOf(delete));
                actions.moveToElement(delete).click().perform();
                sleep(500);
                
                try {
                    Alert alert = driver.switchTo().alert();
                  
                    if(i%2==0) {
                    	alert.dismiss();	
                    }
                    else {
                    	  alert.accept();
                          ok.click();
                    }
                    System.out.println("Alert was present and accepted.");
                } catch (NoAlertPresentException e) {
                    System.out.println("No alert was present.");
                    actions.moveToElement(searchBar).click().perform();
                    clear.click();
                    robot.keyPress(KeyEvent.VK_ENTER);
                    robot.keyRelease(KeyEvent.VK_ENTER);
                }

                sleep(1000);
                checkbox1.click();
                new Select(Status).selectByIndex(2);
                Apply.click();
                wait.until(ExpectedConditions.visibilityOf(ok));
                ok.click();
                
                checkboxmulti.click();
                new Select(Status).selectByIndex(1);
                Apply.click();
                wait.until(ExpectedConditions.visibilityOf(ok));
                ok.click();
             
                System.out.println(i+" iteration completed");
                
            	 fileOut = new FileOutputStream("D://KSP_Admin//ksp_automation.xlsx");
			    workbook.write(fileOut);
            }

         

        } catch (Exception e) {
            e.printStackTrace();
        }
        
        finally {
            try {
                if (fileOut != null) {
                    fileOut.close();
                }
                if (fis != null) {
                    fis.close();
                }
                
            } catch (IOException e) {
                e.printStackTrace();
            }}
    }

    private String getAttributeValue(By locator) {
         return wait.until(ExpectedConditions.presenceOfElementLocated(locator)).getAttribute("value");
    }
} 
