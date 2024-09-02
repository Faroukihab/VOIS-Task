package tests;

import org.testng.annotations.Test;

import pages.sendingEmail;

public class testSendingmail extends TestBase {
	
	sendingEmail send;
	String filePath = "C:\\Users\\fihab\\Downloads\\New folder\\TaskData.xlsx";
	@Test
	public void excecutesendingmail() throws InterruptedException
   {
		send = new sendingEmail(driver);
	   send.openmgmail();
	   Thread.sleep(1000);
	   send.sendEmail();
	   send.attachFile(filePath);
	  
   }
}
