package com.excers.selenium.training;

import org.testng.annotations.Test;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;

public 	class PhPCollabTest {
  	XSSFRow row;
    File file ;
    FileInputStream fis;
    String CurrentSheet;
    XSSFWorkbook workbook;
    XSSFSheet spreadsheet;
    Iterator < Row >  rowIterator;
    WebDriver driver;
    
    
@Test(priority=1)
public void CreateUser() throws IOException, InterruptedException {
	driver.findElement(By.linkText("User Management")).click();
	
	row = (XSSFRow) rowIterator.next();
	while (rowIterator.hasNext()) {
	    row = (XSSFRow) rowIterator.next();
	    driver.findElement(By.xpath("//img[@alt='Add']")).click();
	    Iterator < Cell >  cellIterator = row.cellIterator();
	    Cell cell = cellIterator.next();
		driver.findElement(By.xpath("//input[@name='un']")).sendKeys(cell.getStringCellValue());
		cell = cellIterator.next();
		driver.findElement(By.xpath("//input[@name='fn']")).sendKeys(cell.getStringCellValue());
		cell = cellIterator.next();
		driver.findElement(By.xpath("//input[@name='pw']")).sendKeys(cell.getStringCellValue());
		driver.findElement(By.xpath("//input[@name='pwa']")).sendKeys(cell.getStringCellValue());
		cell = cellIterator.next();
		List<WebElement> permissions=driver.findElements(By.xpath("//input[@name='perm']"));
		  
		if(cell.getStringCellValue().equals("Project Manager"))
		{
			permissions.get(0).click();
		} else if (cell.getStringCellValue().equals("Project Administrator")) {
			permissions.get(3).click();
		} else if (cell.getStringCellValue().equals("User")) {
			permissions.get(1).click();
		}
		driver.findElement(By.xpath("//input[@name='Save']")).click();
		Thread.sleep(1000);
		System.out.println(driver.findElement(By.xpath("//td")).getText());
		Assert.assertEquals(driver.findElement(By.xpath("//td")).getText(), "Success : Addition succeeded");
		
	}
	
	
}

@Test(priority=2, dependsOnMethods={"CreateUser"})
  public void CreateProject() throws IOException, InterruptedException {
	row = (XSSFRow) rowIterator.next();
	while (rowIterator.hasNext()) {
		row = (XSSFRow) rowIterator.next();
		driver.findElement(By.linkText("Projects")).click();
		driver.findElement(By.xpath("//img[@name='saP0']")).click();
		Iterator < Cell >  cellIterator = row.cellIterator();			
		Cell cell = cellIterator.next();
		driver.findElement(By.xpath("//input[@name='pn']")).sendKeys(cell.getStringCellValue());
		cell = cellIterator.next();
		Select priority=new Select(driver.findElement(By.xpath("//select[@name='pr']")));
		priority.selectByVisibleText(cell.getStringCellValue());
		
		cell = cellIterator.next();
		Select owner=new Select(driver.findElement(By.xpath("//select[@name='pown']")));
		owner.selectByVisibleText(cell.getStringCellValue());
		driver.findElement(By.xpath("//input[@value='Save']")).click();
		Thread.sleep(2000);
		Assert.assertEquals(driver.findElement(By.xpath("//td")).getText(), "Success : Addition succeeded");
	 }

  }	

@BeforeMethod
public void beforeMethod(Method method) throws IOException {
System.out.println("Here is the magic "+method.getName());
spreadsheet=workbook.getSheet(method.getName());
rowIterator = spreadsheet.iterator();

}

@AfterMethod
public void afterMethod() {
}

@BeforeTest
public void beforeTest() throws IOException {
  file= new File("E:\\Murali Excers Training\\PhpCollab.xlsx");
  fis = new FileInputStream(file);
  workbook = new XSSFWorkbook(fis);		  
  System.setProperty("webdriver.chrome.driver", "F:\\chromedriver\\chromedriver.exe");
  driver=new ChromeDriver();
  driver.manage().window().maximize();
  driver.get("http://localhost/");
  driver.findElement(By.xpath("//input[@name='loginForm']")).sendKeys("admin");
  driver.findElement(By.xpath("//input[@name='passwordForm']")).sendKeys("phpcadmin");
  driver.findElement(By.xpath("//input[@name='save']")).click();		  
  
}

@AfterTest
public void afterTest() throws IOException {
  fis.close();
  driver.findElement(By.linkText("Log Out")).click();
  driver.quit();
  	  
}
  
}
