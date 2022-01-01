package excel_File_Read;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.time.Duration;

import javax.swing.Action;

import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class Calculator {

	public static void main(String[] args) throws IOException, InterruptedException {

		System.setProperty("webdriver.chrome.driver", "C:\\Users\\Mr.Sagar Chaudhari\\OneDrive - Inspira Enterprise India Private Limited\\Desktop\\priya  study meterial\\Selenium\\Selenium Driver\\chromedriver_win32\\chromedriver.exe");
		WebDriver driver=new ChromeDriver();
		driver.get("https://www.investor.gov/financial-tools-calculators/calculators/compound-interest-calculator");
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10) );
		FileInputStream file =new FileInputStream("C:\\Users\\Mr.Sagar Chaudhari\\OneDrive - Inspira Enterprise India Private Limited\\Desktop\\Priya\\Data.xlsx");
		XSSFWorkbook workbook= new XSSFWorkbook(file);
		XSSFSheet sheet=workbook.getSheet("MoneyControl");
		
		int rowcount =sheet.getLastRowNum();
		
		for(int row=1;row<rowcount;row++)
		{
			XSSFRow currentRow =sheet.getRow(row);
			
	int InitialInvestment =(int)currentRow.getCell(0).getNumericCellValue();
	int MonthlyContribution=(int)currentRow.getCell(1).getNumericCellValue();
	int LengthofTimeInYrs=(int)currentRow.getCell(2).getNumericCellValue();
	int EstimatedInterestRate=(int)currentRow.getCell(3).getNumericCellValue();

	
	driver.findElement(By.id("edit-principal")).sendKeys(String.valueOf(InitialInvestment));
	driver.findElement(By.id("edit-addition")).sendKeys(String.valueOf(MonthlyContribution));
	driver.findElement(By.id("edit-num-years")).sendKeys(String.valueOf(LengthofTimeInYrs));
	driver.findElement(By.id("edit-interest-rate")).sendKeys(String.valueOf(EstimatedInterestRate));

	
	Actions act= new Actions(driver);
	act.moveToElement(driver.findElement(By.id("edit-submit"))).click().build().perform();
	
	
   driver.findElement(By.xpath("//*[@id='sidebar']//print-preview-button-strip//div/cr-button[2]")).click();
	
	act.moveToElement(driver.findElement(By.id("edit-reset"))).click().build().perform();

			
		}
		
System.out.println("Data Driven Test is Uccessfully done");
	}

}
