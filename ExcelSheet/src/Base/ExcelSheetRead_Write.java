package Base;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.Assert;


public class ExcelSheetRead_Write {
    
	public String ReadExcel(String SheetName ) throws IOException  {
String data="";
//		try {
//		FileInputStream filepath =new FileInputStream("C:\\ExcelSheet\\ExcelSheetFolder\\ExcelSheet.xlsx");
//		XSSFWorkbook Wb =new XSSFWorkbook(filepath);
//		XSSFSheet sheet = Wb.getSheet(SheetName);
//		Row r =sheet.getRow(rNo);
//		Cell c =r.getCell(cNo);
//		data=c.getStringCellValue();
//		}
//		catch(Exception e)
//		{
//			e.printStackTrace();
//		}
//		return data;
		FileInputStream filepath =new FileInputStream("C:\\ExcelSheet\\ExcelSheetFolder\\ExcelSheet.xlsx");
				XSSFWorkbook Wb =new XSSFWorkbook(filepath);
	XSSFSheet sheet = Wb.getSheet(SheetName);
	        int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
//	        Row row = sheet.getRow(0);
//	        data = new Object[rowCount][row.getLastCellNum()];
	        for (int i = 1; i <= rowCount; i++) {
	          Row  row = sheet.getRow(i);
	            for (int j = 0; j < row.getLastCellNum(); j++) {
//	                data[i - 1][j] = getCellData(row.getCell(j).getCellType(), row, j);
	            	System.out.print(row.getCell(j).getStringCellValue()+"|| ");
	            }
	        }
	        return data;
	}
	
	public static void main(String[] args) throws InterruptedException, IOException {
		ExcelSheetRead_Write obj =new ExcelSheetRead_Write();
		String Us=obj.ReadExcel("Sheet1");
		//System.out.println("usename: "+Us);
		String Ps=obj.ReadExcel("Sheet1");
		//System.out.println("password: "+Ps);
		  obj.WriteExcel("sheet1", 2, 3, "pass");
		
		 ChromeOptions option = new ChromeOptions();
		  option.addArguments("--remote-allow-origins=*");
	      WebDriver driver=new ChromeDriver(option);
		driver.get("https://demoqa.com/login");
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		
		  WebElement UserName = driver.findElement(By.xpath("//input[@id='userName']"));
		  WebElement Password = driver.findElement(By.xpath("//input[@id='password']"));
		  WebElement loginBtn = driver.findElement(By.xpath("//button[@id='login']"));
		  UserName.sendKeys(Us);
		  Password.sendKeys(Ps);
		  loginBtn.click();
		  Thread.sleep(2000);
		  String ExpectedUrl = "https://demoqa.com/profile";
		  String ActualUrl =driver.getCurrentUrl();
		  Assert.assertEquals(ActualUrl, ExpectedUrl);
		  System.out.println("done");
		
		}

	public void WriteExcel(String SheetName , int rNo,int cNo,String data) 
	{
		try {
			FileInputStream filepath =new FileInputStream("C:\\ExcelSheet\\ExcelSheetFolder\\ExcelSheet.xlsx");
		XSSFWorkbook wb =new XSSFWorkbook(filepath);
		XSSFSheet sheet =wb.getSheet(SheetName);
		Row r=sheet.getRow(rNo); 
		     Cell c =r.createCell(cNo);
		     c.setCellValue(data);
		     FileOutputStream fo =new FileOutputStream("C:\\ExcelSheet\\ExcelSheetFolder\\ExcelSheet.xlsx"); 
		wb.write(fo);
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}
	
}
