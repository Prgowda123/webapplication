package trial;

import java.net.HttpURLConnection;
import java.net.URL;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class Broken_link_finder {
@Test
public void sample()
{
	ChromeDriver driver = new ChromeDriver();
	driver.manage().window().maximize();
	driver.get("https://cpc454.ksp-recruitment.in/");
	List<WebElement> alllink = driver.findElements(By.tagName("a"));
	System.out.println("Total link found :" + alllink.size());
	
	for (WebElement Element : alllink) {
		String url = Element.getAttribute("href");
		
		if(url==null || url.isEmpty())
		{
			System.out.println("⚠️ Skipped: Empty or null URL");
            continue;
		}
		

        try {
            HttpURLConnection connection = (HttpURLConnection) (new URL(url).openConnection());
            connection.setRequestMethod("HEAD");
            connection.connect();
            int responseCode = connection.getResponseCode();

            if (responseCode >= 400) {
                System.out.println("❌ Broken link: " + url + " | Status: " + responseCode);
            } else {
                System.out.println("✅ Valid link: " + url + " | Status: " + responseCode);
            }

        } catch (Exception e) {
            System.out.println("⚠️ Exception checking link: " + url);
        }
	}
	driver.quit();
}
}
