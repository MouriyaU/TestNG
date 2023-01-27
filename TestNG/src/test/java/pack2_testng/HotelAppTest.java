package pack2_testng;

import org.testng.annotations.Test;
import org.testng.annotations.Test;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Parameters;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.testng.annotations.AfterClass;

public class HotelAppTest {
	  WebDriver driver;

	
  @Test
  public void function() {
	  WebElement uName = driver.findElement(By.id("username"));

	  uName.sendKeys("santosh13");
	  // Passing the password


	  driver.findElement(By.id("password")).sendKeys("Snoopy1234");

	  driver.findElement(By.cssSelector("input.login_button[value=Login]")).click();
  }

  @Parameters("browser") 
  @BeforeClass
  public void beforeClass(String browser) {
	  if(browser.equalsIgnoreCase("edge"))
	  {
		  System.setProperty("webdriver.edge.driver", "C:\\Drivers\\msedgedriver.exe");
		  driver = new EdgeDriver();
	  }
	  else if(browser.equalsIgnoreCase("chrome"))
	  {
		  System.setProperty("webdriver.chrome.driver", "C:\\Drivers\\chromedriver.exe");
		  driver = new ChromeDriver();
	  }
	  driver.get("http://adactinhotelapp.com");
  }
  

  @AfterClass
  public void afterClass() {
	  driver.quit();
  }

}
