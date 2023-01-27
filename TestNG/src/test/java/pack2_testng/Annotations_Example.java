package pack2_testng;

import org.testng.annotations.Test;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeClass;
import org.testng.Reporter;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.AfterSuite;

public class Annotations_Example {
	@Test(priority=-1) 
	// order of the execution will be based on the alphebetical order , also on ACII inputs given in function name 
public void afunction2() {
	  System.out.println("Function 2 output ");
	  Reporter.log("I am inside function 2 ");
  }
	
  @Test(priority=1)
  public void Function1() {
	  System.out.println("Sample text 9 : function 1 ");
	  Reporter.log("I am inside function 1 ");
  }
  
  @Test(enabled=false)
  public void Function3() {
	  System.out.println("Sample text 10 : function 3 ");
	  Reporter.log("I am inside function 3 ");
  }
  
  @BeforeMethod
  public void beforeMethod() {
	  System.out.println("Sample text 1 : Before Method");
  }

  @AfterMethod
  public void afterMethod() {
	  
	  System.out.println("Sample text 2 : After method  ");
  }

  @BeforeClass
  public void beforeClass() {
	  System.out.println("Sample text 3 : Before Class ");
  }

  @AfterClass
  public void afterClass() {
	  System.out.println("Sample text 4 : After class  ");
  }

  @BeforeTest
  public void beforeTest() {
	  System.out.println("Sample text 5 : Before test  ");
  }

  @AfterTest
  public void afterTest() {
	  System.out.println("Sample text 6 : After test ");
  }

  @BeforeSuite
  public void beforeSuite() {
	  System.out.println("Sample text 7 : Before Suite ");
  }
  

  @AfterSuite
  public void afterSuite() {
	  System.out.println("Sample text 8 ");
  }

}
