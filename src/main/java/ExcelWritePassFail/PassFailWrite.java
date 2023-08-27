package ExcelWritePassFail;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class PassFailWrite {
	static WebDriver driver;
	public static void main(String[] args) throws IOException, InterruptedException {
			
		System.setProperty("webdriver.chrome.driver","C:\\browserdrivers\\chromedriver.exe");
		
		
		
		File f=new File("C:\\Users\\biswa\\Downloads\\kanha111.xlsx");
		FileInputStream fio=new FileInputStream(f);
		XSSFWorkbook workbook=new XSSFWorkbook(fio);
		XSSFSheet sheet=workbook.getSheet("sheet1");
	int rowCount=sheet.getPhysicalNumberOfRows();
	int colCount=sheet.getRow(0).getPhysicalNumberOfCells();
	System.out.println(rowCount);
	System.out.println(colCount);
	
	XSSFRow row;
	XSSFCell cell;
	String username=null;
	String password=null;
	
	
	for(int i=0;i<rowCount;i++)
	{
		row=sheet.getRow(i);
		for(int j=0;j<colCount;j++) {
			cell=row.getCell(j);
			if(j==0)
			{
				username=cell.getStringCellValue();
				System.out.println(username);
			}
			else {
				password=cell.getStringCellValue();
				System.out.println(password);
			}
		}
			
			driver=new ChromeDriver();
			driver.manage().window().maximize();
			driver.get("https://mail.apmosys.com/webmail/");
			Thread.sleep(3000);
			driver.findElement(By.name("email-address")).sendKeys(username);
			driver.findElement(By.name("next")).click();
			Thread.sleep(2000);
			driver.findElement(By.name("password")).sendKeys(password);
			driver.findElement(By.name("remember-me")).click();
			Thread.sleep(3000);
			driver.findElement(By.name("next")).click();
			Thread.sleep(9000);
			
			String Result=null;
			try {
				Boolean check=driver.findElement(By.id("gui.frm_main.bar.tree.add_container.btn_add#main")).isDisplayed();
				if(check==true)
				{
					Result="PASS";
					cell=row.createCell(3);
					cell.setCellValue(Result);
				}
				System.out.println("username===="+username+"   password===="+password+"   Result==="+Result);
			}
			catch(Exception e) {
				Boolean check1=driver.findElement(By.className("o-well__text")).isDisplayed();
				if(check1==true)
				{
					Result="FAIL";
					cell=row.createCell(3);
					cell.setCellValue(Result);
				}
				System.out.println("username===="+username+"   password===="+password+"   Result==="+Result);
			
				
			}
			
			
			
		
		
	}
	
	FileOutputStream fos=new FileOutputStream(f);
	workbook.write(fos);
	fos.close();

	}

}
