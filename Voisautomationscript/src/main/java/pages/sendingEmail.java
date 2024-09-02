package pages;
import java.util.NoSuchElementException;

import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class sendingEmail extends PageBase {

  
    @FindBy(xpath = "//div[@class='T-I T-I-KE L3']")
    private WebElement composeBtn;
    
    @FindBy(id =  ":bb")
    private WebElement sendToField;
    
    @FindBy(name  =  "subjectbox")
    private WebElement subjectField;
    
    @FindBy(id  =  ":8w")
    private WebElement mailboady;
    
    @FindBy(id  =  ":7c")
    private WebElement sendingBtn;
    
    @FindBy(id  =  ":9a")
    private WebElement attachmentBtn;
    
    @FindBy(id  =  "identifierId")
    private WebElement emailfield;
    
    @FindBy(name  =  "Passwd")
    private WebElement passfield;
    
    @FindBy(css =  "input[type='file']")
    private WebElement fileInput;
    
    @FindBy(xpath =  "//span[@jsname='V67aGc' and contains(@class, 'VfPpkd-vQzf8d')]")
    private WebElement nextBtn;

    

    // Constructor
    public sendingEmail(WebDriver driver) {
        super(driver);
        
        PageFactory.initElements(driver, this);
    }

    
    
    public void openmgmail() throws InterruptedException
    {
    	emailfield.sendKeys("roa522897@gmail.com");
    	Thread.sleep(1000);
    	emailfield.sendKeys(Keys.ENTER);
    	passfield.sendKeys("3101967Sawsan");
    	passfield.sendKeys(Keys.ENTER);
    	Thread.sleep(1000);
    	
    }
    public void sendEmail() 
    {
    	
    	composeBtn.click();
    	sendToField.sendKeys("islam.hassan@vodafone.com");
    	subjectField.sendKeys("Vois Task");
    	mailboady.sendKeys("Dear Eng. Islam,\n\nI hope this email finds you well. I'm sending this to complete the task requirements.\n\nBest regards,\nFarouk Ihab");
    	
    	
    }
    public void attachFile(String filePath) throws InterruptedException {
    
        attachmentBtn.click();
        fileInput.sendKeys(filePath);
        Thread.sleep(2000); 
        sendingBtn.click();
        Thread.sleep(2000); 
    }

}
