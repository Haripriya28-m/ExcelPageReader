package ExcelPageReaderFile.ExcelPageReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class SampleExcelReader {
	
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
    public static void verifyExcelReader() throws IOException
    {
    File excelFile = new File("D:\\Book1.xlsx");
    
    FileInputStream file1 = new FileInputStream(excelFile);
    
    XSSFWorkbook workbook = new XSSFWorkbook(file1);
    
    XSSFSheet sheet = workbook.getSheetAt(0);
/*    
    XSSFRow row1=sheet.getRow(0);
    XSSFCell cell1=row1.getCell(0);
    
    XSSFRow row2=sheet.getRow(1);
    XSSFCell cell2=row2.getCell(0);
    
    String text1 = cell1.getStringCellValue();
    System.out.println("text1 value:"+ text1);
    
    WebElement element = driver.findElement(By.name("user-name"));
	element.sendKeys(cell1.getStringCellValue());
    
    String text2 = cell2.getStringCellValue();
    System.out.println("text2 value:"+ text2);
    
    WebElement element1 = driver.findElement(By.name("password"));
	element1.sendKeys(cell2.getStringCellValue());
    
	
	driver.findElement(By.name("login-button")).click();
*/	
	WebElement username=driver.findElement(By.name("user-name"));
	WebElement password=driver.findElement(By.name("password"));
	
	int rowcount = 0;
	for(int i = 0;i <= rowcount; i++)
	{
		username.sendKeys(sheet.getRow(i).getCell(0).getStringCellValue());
		password.sendKeys(sheet.getRow(i).getCell(1).getStringCellValue());
		
		driver.findElement(By.name("login-button")).click();
	}
    
/*    for(Row row:sheet) {
    	for(Cell cell:row) {
    		System.out.println(cell.getStringCellValue());
    		WebElement element = driver.findElement(By.name("user-name"));
    		element.sendKeys(cell.getStringCellValue());
    		
    		WebElement element1 = driver.findElement(By.name("password"));
    		element1.sendKeys(cell.getStringCellValue());
    	}
    }
  */  
    workbook.close();
    file1.close();
  
	}
}




