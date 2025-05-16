package trial;

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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class Excel {
    

    @Test
    public void runAutomation() throws IOException, InterruptedException, AWTException {
        try {
            WebDriver driver = new ChromeDriver();
            driver.manage().window().maximize();
            driver.get("https://ksp-recruitment.in/");
            Thread.sleep(4000);
            JavascriptExecutor jss = (JavascriptExecutor) driver;
            jss.executeScript("window.scrollBy(0,500)", "");
            Thread.sleep(4000);
            driver.findElement(By.linkText("CLICK HERE TO GO TO RECRUITMENT - 2022")).click();
            Thread.sleep(4000);

            Set<String> allwh = driver.getWindowHandles();
            for (String wh : allwh) {
                driver.switchTo().window(wh);
            }
            Thread.sleep(4000);
            jss.executeScript("window.scrollBy(0,1500)", "");
            Thread.sleep(4000);
            driver.findElement(By.xpath("//td[text()=' Armed Police Constable (Male & Male Transgender) (CAR/DAR)-2022 ']/following-sibling::td[5]")).click();
            Thread.sleep(4000);
            Set<String> allwh1 = driver.getWindowHandles();
            ArrayList<String> lan = new ArrayList<String>(allwh1);
            driver.switchTo().window(lan.get(2));
            Thread.sleep(4000);

            String parentWindow = driver.getWindowHandle();
            Set<String> allWindows = driver.getWindowHandles();
            allWindows.remove(parentWindow);
            for (String allw : allWindows) {
                driver.switchTo().window(allw);
                driver.close();
            }
            driver.switchTo().window(parentWindow);

            driver.findElement(By.xpath("//a[@class='nav-menu nav-myapp-menu login-btn p-2']")).click();
            Thread.sleep(4000);

            String mainWindowHandle = driver.getWindowHandle();
            Set<String> windowHandles = driver.getWindowHandles();
            for (String handle : windowHandles) {
                if (!handle.equals(mainWindowHandle)) {
                    driver.switchTo().window(handle);
                    break;
                }
            }

            FileInputStream fis = new FileInputStream("D://Automation_data//Apc_3064.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);
            int rowcount = sheet.getPhysicalNumberOfRows();

            for (int i = 1; i < 2; i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    String applicationNo = getCellValue(row.getCell(0));
                    String dob = getCellValue(row.getCell(1));

                    System.out.println("Application No: " + applicationNo);
                    System.out.println("DOB: " + dob);

                    driver.findElement(By.id("Login_ApplicantId")).sendKeys(applicationNo);
                    Thread.sleep(1000);
                    driver.findElement(By.id("Login_DateOfBirth")).sendKeys(dob);
                    Thread.sleep(1000);

                    Robot r = new Robot();
                    r.keyPress(KeyEvent.VK_ENTER);
                    r.keyRelease(KeyEvent.VK_ENTER);
                    Thread.sleep(2000);

                    jss.executeScript("window.scrollBy(0,500)", "");
                    Thread.sleep(2000);
                    WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(2));
                    Sheet sheet1 = workbook.getSheetAt(1);
                    Row row1 = sheet1.createRow(sheet1.getPhysicalNumberOfRows());

                    WebElement appno = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//td[contains(text(),' Application No. ')]//../td[2]")));
                    row1.createCell(0).setCellValue(appno.getText());

                    WebElement DOB = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//td[contains(text(),' Date of Birth')]//../td[2]")));
                    row1.createCell(1).setCellValue(DOB.getText());

                    WebElement written = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h5[contains(text(),'Written Exam')]")));
                    if (isElementClickable(driver, written)) {
                        WebElement examdate = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//td[contains(text(),'WRITTEN EXAM DATE' )]/following-sibling::td[1]")));
                        row1.createCell(2).setCellValue(examdate.getText());
                    } else {
                        row1.createCell(2).setCellValue("WRITTEN EXAM IS NOT COMPLETED");
                    }

                    try {
                        WebElement ENDURANCE = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h5[contains(text(),'ENDURANCE TEST & PHYSICAL STANDARD TEST' )]")));
                        if (isElementClickable(driver, ENDURANCE)) {
                            WebElement ENDURANCETEST = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h5[contains(text(),'ENDURANCE TEST & PHYSICAL STANDARD TEST' )]/../..//div/table/tbody/tr[1]/th/following-sibling::td[2]")));
                            row1.createCell(3).setCellValue(ENDURANCETEST.getText());
                        }
                    } catch (Exception e) {
                        System.out.println("ETPST is COMPLETE");
                    }

                    try {
                        WebElement MEDICAL = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h5[contains(text(),' MEDICAL EXAM' )]")));
                        if (isElementClickable(driver, MEDICAL)) {
                            WebElement DOC = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//td[text()=' DOCUMENT VERIFICATION DATE & TIME. ']/following-sibling::td[1]")));
                            row1.createCell(4).setCellValue(DOC.getText());
                        } else {
                            row1.createCell(4).setCellValue("DV not COMPLETED");
                        }
                    } catch (Exception e) {
                        System.out.println("DV is not complete");
                    }

                    try {
                        WebElement MED = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//td[text()=' MEDICAL EXAM DATE & TIME. ']/following-sibling::td[1]")));
                        row1.createCell(5).setCellValue(MED.getText());
                    } catch (Exception e) {
                        row1.createCell(5).setCellValue("MV not COMPLETED");
                    }

                    try {
                        String st = "Office of the Superintendent Of Police, BIDAR District";
                        WebElement ele = driver.findElement(By.xpath("//td[text()=' DOCUMENT VERIFICATION VENUE. ']/following-sibling::td[1]"));
                        if (st.equals(ele.getText())) {
                            System.out.println("Both the venues are same");
                        }
                    } catch (Exception e) {
                    }

                    driver.close();
                    Thread.sleep(2000);
                    driver.switchTo().window(mainWindowHandle);
                    driver.findElement(By.xpath("//a[@class='nav-menu nav-myapp-menu login-btn p-2']")).click();

                    Set<String> windowHandles1 = driver.getWindowHandles();
                    for (String handle1 : windowHandles1) {
                    	
                        if (!handle1.equals(mainWindowHandle)) {
                            driver.switchTo().window(handle1);
                            break;
                        }
                    }

                    FileOutputStream file = new FileOutputStream("D://Automation_data//Apc_3064.xlsx");
                    workbook.write(file);
                    file.close();
                }}
            
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
                
            

            workbook.close();
            fis.close();
            driver.quit();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    return new SimpleDateFormat("dd-MM-yyyy").format(date);
                } else {
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
