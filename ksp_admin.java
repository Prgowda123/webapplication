package trial;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.text.ParseException;
import java.time.Duration;
import java.util.Set;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class ksp_admin {

	@Test
	public void sample1() throws InterruptedException, AWTException, ParseException
	{
		ChromeDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("http://172.10.1.159:9052/Masters/ApplyingType");
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
	
		Robot r = new Robot();
		JavascriptExecutor jss = (JavascriptExecutor) driver;
		Actions act = new Actions(driver);
//		WebElement view = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[contains(text(),'visibility')]")));
//		act.moveToElement(view).click().perform();
//
//		
//		switchToNewWindow(driver);
//		Thread.sleep(2000);
//		
		WebElement viewcode = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='code']")));
		act.moveToElement(viewcode).doubleClick().perform();
		wait.until(ExpectedConditions.visibilityOf(viewcode));
		wait.until(ExpectedConditions.elementToBeClickable(viewcode));
		String getviewcode = viewcode.getAttribute("value");	
		System.out.println(getviewcode);
		
		WebElement viewtitle = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='title']")));
		String getviewtitle = viewtitle.getAttribute("value");	
		System.out.println(getviewtitle);
		
		WebElement viewoi = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='orderIndex']")));
		String getviewoi = viewoi.getAttribute("value");	
		System.out.println(getviewoi);
		
    	WebElement viewstatus = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='status']")));
		String getviewstatus = viewstatus.getAttribute("value");	
		System.out.println(getviewstatus);
		 wait.until(ExpectedConditions.presenceOfElementLocated(By.linkText("arrow_back"))).click();
		 

	    	switchToNewWindow(driver);
	    	
//	    	WebElement search = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("search-bar")));	
//			act.moveToElement(search).click().perform();	
//			r.keyPress(KeyEvent.VK_BACK_SPACE);
//			Thread.sleep(2000);
//			r.keyRelease(KeyEvent.VK_BACK_SPACE);
//			
	    	WebElement edit = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("(//span[contains(text(),'edit')])[1]")));
			act.moveToElement(edit).click().perform();

			switchToNewWindow(driver);
			Thread.sleep(2000);
}
	
	private void switchToNewWindow(WebDriver driver) {
		Set<String> windowHandles = driver.getWindowHandles();
		for (String windowHandle : windowHandles) {
		    driver.switchTo().window(windowHandle);
		}}
}