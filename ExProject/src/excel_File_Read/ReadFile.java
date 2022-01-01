package excel_File_Read;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFile {

	public static void main(String[] args) throws IOException {

		FileInputStream file =new FileInputStream("C:\\Users\\Mr.Sagar Chaudhari\\OneDrive - Inspira Enterprise India Private Limited\\Desktop\\priya  study meterial\\Selenium\\Selenium Read Data.xlsx");
		XSSFWorkbook workbook =new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
		
		
		int rowcount =sheet.getLastRowNum();
		int colcount =sheet.getRow(0).getLastCellNum();
		
		for(int i=0;i<rowcount;i++)
		{
		     XSSFRow CurrentRow=sheet.getRow(i);
		     
		     for(int j=0;j<colcount;j++)
		     {
		    	 String Value=CurrentRow.getCell(j).toString();
		    	 System.out.print("      " +Value);
		    	 
		     }
		     System.out.println();
		       
		     
		}
		

	}

}
