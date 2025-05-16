package psi_sports;

import java.awt.AWTException;
import java.awt.Robot;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.Set;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.*;

import org.testng.annotations.Test;

public class Calender {
    @Test
    public void sample() throws InterruptedException, AWTException, IOException {
        ChromeDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.get("http://172.10.1.159:9033");
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));


        Actions act = new Actions(driver);
        Robot r = new Robot();
       

        // Load Excel File
        FileInputStream fis1 = new FileInputStream("D://Automation_data//Calender.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis1);
        Sheet sheet = workbook.getSheetAt(0);

        int rowCount = sheet.getPhysicalNumberOfRows();

        // Loop through Excel rows
        for (int i = 1; i < rowCount; i++) { // Start from 1 to skip header
            Row row = sheet.getRow(i);

            if (row == null) {
                System.out.println("Skipping empty row: " + i);
                continue;
            }

            Cell cell = row.getCell(0); // Assuming DOB is in column A (index 0)
            if (cell == null || cell.getCellType() == CellType.BLANK) {
                System.out.println("Skipping empty cell in row: " + i);
                continue;
            }

            // Read Date from Excel (Handle both String and Date formats)
            String date = "";
            if (cell.getCellType() == CellType.STRING) {
                date = cell.getStringCellValue().trim();
            } else if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                SimpleDateFormat sdf = new SimpleDateFormat("dd-MMMM-yyyy");
                date = sdf.format(cell.getDateCellValue());
            }

            if (date.isEmpty()) {
                System.out.println("Skipping row due to empty date: " + i);
                continue;
            }

            System.out.println("Selecting date: " + date);

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

            Thread.sleep(1000);
            jss.executeScript("window.scrollBy(0,100)", "");
            
            jss.executeScript("window.scrollBy(0,1200)", "");
            Thread.sleep(1000);
            // Split date into day, month, year
            String[] dateParts = date.split("-");
            String day = dateParts[0];
            String month = dateParts[1];
            String year = dateParts[2];
            day = String.valueOf(Integer.parseInt(day));
            selectDate(driver, day, month, year);
            Thread.sleep(2000);
            WebElement datePicker = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("(//input[@class='form-control form-control input' and @type='text' and @readonly='readonly'])[1]")));
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", datePicker);

            // Select the year
            WebElement yearInput = driver.findElement(By.xpath("(//input[@class='numInput cur-year'])[1]"));
            yearInput.clear();
            Thread.sleep(1000);
            yearInput.sendKeys(year);
            Thread.sleep(1000);
            yearInput.sendKeys(Keys.ENTER);
            Thread.sleep(1000);

            // Select the month from dropdown
            WebElement monthDropdown = driver.findElement(By.xpath("(//select[contains(@class,'flatpickr-monthDropdown-months')])[1]"));
            Thread.sleep(1000);
            monthDropdown.click();
            Select monthSelect = new Select(monthDropdown);
            monthSelect.selectByVisibleText(month);
            Thread.sleep(1000);
        
            // Select the day
            WebElement dayElement = driver.findElement(By.xpath("//span[text()='" + day + "']"));
            dayElement.click();
            
           // driver.close();
        }

        workbook.close();
        fis1.close();
      //  driver.quit();
    }

    public static void selectDate(WebDriver driver, String day, String month, String year) {
        // Method stub for selecting a date
    }
}
