package ExcelPageReaderFile.ExcelPageReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class ExcelReaderSwags {

	
	static WebDriver driver;
	@BeforeMethod
	public static  void setup()
	{
		WebDriverManager.chromedriver().setup();
		driver=new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(30,TimeUnit.SECONDS);
		driver.get("https://www.saucedemo.com/");
	}
	
	 @Test
	    public static void verifyExcelReader1() throws IOException
	    {
	    File excelFile = new File("D:\\Book2.xlsx");
	    FileInputStream file1 = new FileInputStream(excelFile);
	    XSSFWorkbook workbook = new XSSFWorkbook(file1);
	    XSSFSheet sheet = workbook.getSheetAt(0);
	    WebElement username=driver.findElement(By.name("user-name"));
		WebElement password=driver.findElement(By.name("password"));
		
		int rowcount = 0;
		for(int i = 0;i <= rowcount; i++)
		{
			username.sendKeys(sheet.getRow(i).getCell(0).getStringCellValue());
			password.sendKeys(sheet.getRow(i).getCell(1).getStringCellValue());
			driver.findElement(By.name("login-button")).click();
			 
			 TakesScreenshot scr = ((TakesScreenshot) driver);
				File file0 = scr.getScreenshotAs(OutputType.FILE);

				FileUtils.copyFile(file0,new File("D:\\Haripriya eclipse\\ExcelPageReader\\pageReaderScreenshot\\login1.png"));
				System.out.println("screenshot of the LoginPage1  is taken");
		}
		 workbook.close();
		    file1.close();
		driver.close(); 
	    }
	     
	@Test
	 public static void verifyExcelReader2() throws IOException  
	 {
		 File excelFile = new File("D:\\Book2.xlsx");
		    FileInputStream file2 = new FileInputStream(excelFile);
		    XSSFWorkbook workbook = new XSSFWorkbook(file2);
		    XSSFSheet sheet = workbook.getSheetAt(0);
		    WebElement username=driver.findElement(By.name("user-name"));
			WebElement password=driver.findElement(By.name("password"));
			
			int rowcount = 1;
			for(int i = 1;i <= rowcount; i++)
			{
				username.sendKeys(sheet.getRow(i).getCell(0).getStringCellValue());
				password.sendKeys(sheet.getRow(i).getCell(1).getStringCellValue());
				driver.findElement(By.name("login-button")).click();
				 
				 TakesScreenshot scr = ((TakesScreenshot) driver);
					File file21 = scr.getScreenshotAs(OutputType.FILE);

					FileUtils.copyFile(file21,new File("D:\\Haripriya eclipse\\ExcelPageReader\\pageReaderScreenshot\\login2.png"));
					System.out.println("screenshot of the LoginPage2  is taken");
			}
			 workbook.close();
			    file2.close();
			driver.close(); 
	 }
	
	
	@Test
	 public static void verifyExcelReader3() throws IOException  
	 {
		 File excelFile = new File("D:\\Book2.xlsx");
		    FileInputStream file3 = new FileInputStream(excelFile);
		    XSSFWorkbook workbook = new XSSFWorkbook(file3);
		    XSSFSheet sheet = workbook.getSheetAt(0);
		    WebElement username=driver.findElement(By.name("user-name"));
			WebElement password=driver.findElement(By.name("password"));
			
			int rowcount = 2;
			for(int i = 2;i <= rowcount; i++)
			{
				username.sendKeys(sheet.getRow(i).getCell(0).getStringCellValue());
				password.sendKeys(sheet.getRow(i).getCell(1).getStringCellValue());
				driver.findElement(By.name("login-button")).click();
				 
				 TakesScreenshot scr = ((TakesScreenshot) driver);
					File file31 = scr.getScreenshotAs(OutputType.FILE);

					FileUtils.copyFile(file31,new File("D:\\Haripriya eclipse\\ExcelPageReader\\pageReaderScreenshot\\login3.png"));
					System.out.println("screenshot of the LoginPage3  is taken");
			}
			 workbook.close();
			    file3.close();
			driver.close(); 
	 }
	
	
	@Test
	 public static void verifyExcelReader4() throws IOException  
	 {
		 File excelFile = new File("D:\\Book2.xlsx");
		    FileInputStream file4 = new FileInputStream(excelFile);
		    XSSFWorkbook workbook = new XSSFWorkbook(file4);
		    XSSFSheet sheet = workbook.getSheetAt(0);
		    WebElement username=driver.findElement(By.name("user-name"));
			WebElement password=driver.findElement(By.name("password"));
			
			int rowcount = 3;
			for(int i = 3;i <= rowcount; i++)
			{
				username.sendKeys(sheet.getRow(i).getCell(0).getStringCellValue());
				password.sendKeys(sheet.getRow(i).getCell(1).getStringCellValue());
				driver.findElement(By.name("login-button")).click();
				 
				 TakesScreenshot scr = ((TakesScreenshot) driver);
					File file41 = scr.getScreenshotAs(OutputType.FILE);

					FileUtils.copyFile(file41,new File("D:\\Haripriya eclipse\\ExcelPageReader\\pageReaderScreenshot\\login4.png"));
					System.out.println("screenshot of the LoginPage4  is taken");
			}
			 workbook.close();
			    file4.close();
			driver.close(); 
	 }
	
	
	@Test
	 public static void verifyExcelReader5() throws IOException  
	 {
		 File excelFile = new File("D:\\Book2.xlsx");
		    FileInputStream file5 = new FileInputStream(excelFile);
		    XSSFWorkbook workbook = new XSSFWorkbook(file5);
		    XSSFSheet sheet = workbook.getSheetAt(0);
		    WebElement username=driver.findElement(By.name("user-name"));
			WebElement password=driver.findElement(By.name("password"));
			
			int rowcount = 4;
			for(int i = 4;i <= rowcount; i++)
			{
				username.sendKeys(sheet.getRow(i).getCell(0).getStringCellValue());
				password.sendKeys(sheet.getRow(i).getCell(1).getStringCellValue());
				driver.findElement(By.name("login-button")).click();
				 
				 TakesScreenshot scr = ((TakesScreenshot) driver);
					File file51 = scr.getScreenshotAs(OutputType.FILE);

					FileUtils.copyFile(file51,new File("D:\\Haripriya eclipse\\ExcelPageReader\\pageReaderScreenshot\\login5.png"));
					System.out.println("screenshot of the LoginPage5  is taken");
			}
			 workbook.close();
			    file5.close();
			driver.close(); 
	 }
	
	
	@Test
	 public static void verifyExcelReader6() throws IOException  
	 {
		 File excelFile = new File("D:\\Book2.xlsx");
		    FileInputStream file6 = new FileInputStream(excelFile);
		    XSSFWorkbook workbook = new XSSFWorkbook(file6);
		    XSSFSheet sheet = workbook.getSheetAt(0);
		    WebElement username=driver.findElement(By.name("user-name"));
			WebElement password=driver.findElement(By.name("password"));
			
			int rowcount = 5;
			for(int i = 5;i <= rowcount; i++)
			{
				username.sendKeys(sheet.getRow(i).getCell(0).getStringCellValue());
				password.sendKeys(sheet.getRow(i).getCell(1).getStringCellValue());
				driver.findElement(By.name("login-button")).click();
				 
				 TakesScreenshot scr = ((TakesScreenshot) driver);
					File file61 = scr.getScreenshotAs(OutputType.FILE);

					FileUtils.copyFile(file61,new File("D:\\Haripriya eclipse\\ExcelPageReader\\pageReaderScreenshot\\login6.png"));
					System.out.println("screenshot of the LoginPage6  is taken");
			}
			 workbook.close();
			    file6.close();
			driver.close(); 
	 }
	
	@AfterMethod
	public static void vsetup()
	{
		driver.quit();
	}
		}

