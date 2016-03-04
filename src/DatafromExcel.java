import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;


public class DatafromExcel {

	public static void main(String[] args)throws EncryptedDocumentException, InvalidFormatException, IOException  {
		// TODO Auto-generated method stub

 	 	WebDriver driver = new FirefoxDriver();
 	 	driver.get("http://localhost/login.do");
 	 
			FileInputStream fis = new FileInputStream("Z:\\actitime.xls");
			Workbook workbook =  WorkbookFactory.create(fis);
			Sheet sheet0 = workbook.getSheetAt(0);
			for(int i= 1;i<=3;i++){
				String username = sheet0.getRow(i).getCell(0).getStringCellValue();
				System.out.println(username);
				driver.findElement(By.xpath(".//*[@id='username']")).sendKeys(username);
			
					String password = sheet0.getRow(i).getCell(1).getStringCellValue();
					System.out.println(password);
					driver.findElement(By.xpath(".//*[@id='loginFormContainer']/tbody/tr[1]/td/table/tbody/tr[2]/td/input")).sendKeys(password);
				    driver.findElement(By.xpath(".//*[@id='loginButton']/div")).click();
				    System.out.println("actitime page opened");
				    driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
					driver.findElement(By.xpath(".//*[@id='logoutLink']")).click();
					
				
	}

}
}