package ksp_admin;

import java.awt.AWTException;

import org.testng.annotations.Test;

public class Runner extends Base_class{
	@Test(enabled = false)
	public void apptype() throws AWTException, InterruptedException
	{
		Pom_class p = new Pom_class(driver);
		p.Master();
		p.apptype();
		p.switc();
		p.add();
		p.switc();
		Thread.sleep(1000);
		p.adddetail();
	}
	
	@Test
	public void sample() throws AWTException
	{
		pom p1=new pom(driver);
		p1.openMasters();
		p1.openApplyingTypes();
		p1.switchWindow();
//		p1.clickAdd();
//		p1.switchWindow();
		p1.addd();
	}

}
