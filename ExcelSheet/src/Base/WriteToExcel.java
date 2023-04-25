package Base;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
public class WriteToExcel {
	public static void main(String[] args) throws IOException, InterruptedException {
	        File file =    new File("C:\\ExcelSheet\\ExcelSheetFolder\\ExcelSheet.xlsx");
	        FileInputStream inputStream = new FileInputStream(file);
	        XSSFWorkbook wb=new XSSFWorkbook(inputStream);
	       XSSFSheet sheet=wb.getSheet("Sheet1");
	        int rowCount=sheet.getLastRowNum()-sheet.getFirstRowNum();
	        System.out.println(rowCount);
	        ChromeOptions option = new ChromeOptions();
	        option.addArguments("--remote-allow-origins=*");
		    WebDriver driver=new ChromeDriver(option);
		    driver.manage().window().maximize();
		    driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
	        driver.get("https://demoqa.com/login");
	     
	        for(int i=1;i<=rowCount;i++) {
	        	 Thread.sleep(2000);
	        	 System.out.println(i);
	        	 WebElement UserName = driver.findElement(By.id("userName"));
				  WebElement Password = driver.findElement(By.xpath("//input[@id='password']"));
				  WebElement loginBtn = driver.findElement(By.xpath("//button[@id='login']"));
	           
				  UserName.clear();
				  Password.clear();
	        	UserName.sendKeys(sheet.getRow(i).getCell(0).getStringCellValue());
	        	 Password.sendKeys(sheet.getRow(i).getCell(1).getStringCellValue());
	            loginBtn.click();
	            Thread.sleep(2000);
	            String ExpectedUrl = "https://demoqa.com/profile";
	      	  String ActualUrl =driver.getCurrentUrl();
	      
	        
	            if (ActualUrl.equalsIgnoreCase(ExpectedUrl)) {
	                
	            	 XSSFCell cell = sheet.getRow(i).createCell(3);
	                cell.setCellValue("PASS");
	                FileOutputStream outputStream = new FileOutputStream("C:\\ExcelSheet\\ExcelSheetFolder\\ExcelSheet.xlsx");
		            wb.write(outputStream);
		            WebElement closebtn = driver.findElement(By.xpath("//button[text()='Log out']"));
		            closebtn.click();
	                
	            } else {
	               
	            	 XSSFCell cell = sheet.getRow(i).createCell(3);
	                cell.setCellValue("FAIL");
	                FileOutputStream outputStream = new FileOutputStream("C:\\ExcelSheet\\ExcelSheetFolder\\ExcelSheet.xlsx");
		            wb.write(outputStream);
		         
	            }
		            
	            driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
	           
	        }
	        

	        wb.close();
	        
	    
	        driver.quit();
	        }
	}


