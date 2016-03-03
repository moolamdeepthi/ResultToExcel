import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.testng.Assert;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
public class Testresut {
	public WebDriver driver;
	String baseurl= "https://mail.blr.velankani.com";	
	static Map<String, Object[]> data = new LinkedHashMap<String, Object[]>();;
	static HSSFWorkbook workbook = new HSSFWorkbook();
	static HSSFSheet sheet = workbook.createSheet("Sample sheet");

@BeforeTest
	public static void excelSheetHeading() {
		
		sheet.setColumnWidth((short)1, (short)10000);
		sheet.setColumnWidth((short)2, (short)10000);
		sheet.setColumnWidth((short)3, (short)10000);
		data.put("1", new Object[] {"TestId", "Action", "Expected","Actual","Status"});
		writeToExcel(data);
		
}
		@Test(priority=0)
		public void openWebpage() throws IOException{
			driver = new FirefoxDriver();
			driver.get(baseurl);
			driver.findElement(By.xpath(".//*[@id='horde_user']")).sendKeys("*******");
			driver.findElement(By.xpath(".//*[@id='horde_pass']")).sendKeys("*********");
			driver.findElement(By.id("login-button")).click();
		     String actual = driver.getTitle();
		     System.out.println(actual);
		     String expected = "Horde :: My Portal";
	
			   if(expected.equals(actual)){
			   data.put("2",new Object[] {"1","Navigate to site and login","Site should open and login sucessfully","Site opens and login sucessfully","Pass"});
			   writeToExcel(data);
			   }else{
				   data.put("2",new Object[] {"1","Navigate to site and login","Site should open and login sucessfully","Failed to login","Fail"});
				   writeToExcel(data);
				      
			   }
		}

		@Test(priority=1)
		public void openSettingsPage(){
			driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
			driver.findElement(By.xpath(".//*[@id='horde-navigation']/div[7]/ul/li/div")).click();
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		Actions abc = new Actions(driver);
			WebElement preferences = driver.findElement(By.xpath(".//*[@id='horde-navigation']/div[7]/ul/li/ul/li[1]/div/a"));
		abc.moveToElement(preferences).perform();
			driver.findElement(By.xpath(".//*[@id='horde-navigation']/div[7]/ul/li/ul/li[1]/ul/li[5]/div/a")).click();
			String actual = driver.getTitle();
			System.out.println(actual);
			String expected = "Mail :: User Preferences"; 
			Assert.assertEquals(actual, expected);
			if(expected.equals(actual)){
				data.put("3",new Object[] {"2","Open Settings Page & Navigate to Preferences","PreferencesMail page should be displayed","PreferencesMail page displayed","Pass"});
				writeToExcel(data);
			}else{
				data.put("3",new Object[] {"2","Open Settings Page & Navigate to Preferences","PreferencesMail page should be displayed","PreferencesMail page not displayed","Fail"});
				writeToExcel(data);

			}
		}
		@Test(priority=2)
		public void deleteAndMovingMessages(){
			driver.findElement(By.xpath(".//*[@id='horde-content']/div/div[3]/div/dl/dt[2]/a")).click();
			driver.findElement(By.xpath(".//*[@id='use_trash']")).click();
			driver.findElement(By.xpath(".//*[@id='prefs']/p/input[1]")).click();
			String actual =driver.findElement(By.xpath(".//*[@id='Growler']/div/div[2]")).getText();
			System.out.println(actual);
			String expected = "Your preferences have been updated.";
			Assert.assertEquals(actual, expected);
			if((expected.equals(actual))){
				data.put("4",new Object[]{"3","User can change settings","Settings should changed","Settings changed","pass"});
				writeToExcel(data);
			}
			else{
				data.put("4",new Object[]{"3","User can change settings","Settings should changed","Settings not changed","Fail"});
				writeToExcel(data);
			}
		}
		@Test(priority=3)
		public void logOut(){
			driver.findElement(By.xpath(".//*[@id='horde-logout']/a")).click();
			String actual = driver.getTitle();
			System.out.println(actual);
			String expected= "Horde :: Log in";
			Assert.assertEquals(actual, expected);
			if(expected.equals(actual)){
				data.put("5",new Object[]{"4","Click on logout","You have been logged out should display","You have been logged out displayed","pass"});
				writeToExcel(data);
			}
			else{
				data.put("5",new Object[]{"4","Click on logout","You have been logged out should display ","You have been logged outnot displayed","Fail"});
				writeToExcel(data);
			}

		}
		
		public static void writeToExcel(Map<String, Object[]> data){
			Set<String> keyset = data.keySet();
			int rownum = 0;
			for (String key : keyset) {
				Row row = sheet.createRow(rownum++);
				Object [] objval = data.get(key);
				int cellnum = 0;
				for (Object obj : objval) {
					Cell cell = row.createCell(cellnum++);
					if(obj instanceof Date) 
						cell.setCellValue((Date)obj);
					else if(obj instanceof Boolean)
						cell.setCellValue((Boolean)obj);
					else if(obj instanceof String)
						cell.setCellValue((String)obj);
					else if(obj instanceof Double)
						cell.setCellValue((Double)obj);}}
			try {
				FileOutputStream out =new FileOutputStream(new File("C:\\Users\\deepthi.moolam\\Desktop\\UserSettings.xls"));
				workbook.write(out);
				out.close();
				System.out.println("Excel written successfully..");

			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}

}		
}
