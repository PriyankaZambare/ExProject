package excel_File_Read;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.time.Duration;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class DataDriven_Test {

	public static void main(String[] args) throws IOException, InterruptedException {

		System.setProperty("webdriver.chrome.driver", "C:\\Users\\Mr.Sagar Chaudhari\\OneDrive - Inspira Enterprise India Private Limited\\Desktop\\priya  study meterial\\Selenium\\Selenium Driver\\chromedriver_win32\\chromedriver.exe");
		WebDriver driver=new ChromeDriver();
		driver.get("https://web.schoollog.in/about");
		
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10) );
		FileInputStream file =new FileInputStream("C:\\Users\\Mr.Sagar Chaudhari\\OneDrive - Inspira Enterprise India Private Limited\\Desktop\\Priya\\Data.xlsx");
		XSSFWorkbook workbook= new XSSFWorkbook(file);
		XSSFSheet sheet =workbook.getSheet("SchoolDemo");
		
		int NoOfRows =sheet.getLastRowNum();
		
		for(int row=1;row<NoOfRows;row++)
		{
		    XSSFRow CurrentRow=	sheet.getRow(row);
		 String Name =  CurrentRow.getCell(0).getStringCellValue();
		String FatherName= CurrentRow.getCell(1).getStringCellValue();
		 String Email = CurrentRow.getCell(2).getStringCellValue();
		 String Contact = CurrentRow.getCell(3).getStringCellValue();
		 String msg=CurrentRow.getCell(4).getStringCellValue();
		 
		 
		 driver.findElement(By.name("name")).sendKeys(Name);
		 driver.findElement(By.name("father_name")).sendKeys(FatherName);
		 driver.findElement(By.name("email")).sendKeys(Email);
		 driver.findElement(By.name("phone")).sendKeys(Contact);
		 driver.findElement(By.name("message")).sendKeys("Hii ");
		 driver.findElement(By.xpath("/html/body/div[1]/div/div[5]/div[1]/div[1]/form/button")).click();//	
		 driver.findElement(By.xpath("/html/body/div[3]/div/div[3]/div/button")).click();
		 Thread.sleep(2000);
	//	 Alert alt =driver.switchTo().alert();
	//    alt.accept();
		
			
			
			
			
			
			
			
			
			
			
			
			
	/*		
			if(driver.getPageSource().contains("This is done"))
			{
				System.out.println("Registration complete for " +row+ "record");
			}
			
			else
			{
				System.out.println("Registration fail for \" +row+ \"record");
			}
		 
	//	 driver.close();
		}
		
		*/

	}

}
}