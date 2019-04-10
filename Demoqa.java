package org.data.sample;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class Demoqa {

	public static String getdata(int rowno, int cellno) throws IOException {
		File loc = new File("C:\\Users\\M\\eclipse-1\\MavenClass\\datas\\Book1.xls");

		FileInputStream stream = new FileInputStream(loc);

		Workbook w = new HSSFWorkbook(stream);

		Sheet s = w.getSheet("Data");

		Row r = s.getRow(rowno);

		Cell c = r.getCell(cellno);

		int type = c.getCellType();

		String name = null;
		if (type == 1) {
			name = c.getStringCellValue();
		}

		if (type == 0) {
			if (DateUtil.isCellDateFormatted(c)) {
				name = new SimpleDateFormat("dd/mm/yyyy").format(c.getDateCellValue());
			} else {
				name = String.valueOf((long) c.getNumericCellValue());

			}
		}
		return name;

	}

	public static void select(WebElement element, int i) {
		Select s = new Select(element);
		s.selectByIndex(i);
	}

	public static void submit(WebElement element) {
		element.click();
	}

	public static void main(String[] args) throws IOException, Exception {

		System.setProperty("webdriver.chrome.driver", "C:\\Users\\M\\eclipse-1\\MavenClass\\driver\\chromedriver1.exe");
		WebDriver driver = new ChromeDriver();
		driver.get("http://www.adactin.com/HotelApp/");
		driver.manage().window().maximize();

		WebElement username = driver.findElement(By.id("username"));
		username.sendKeys(getdata(1, 0));

		WebElement password = driver.findElement(By.id("password"));
		password.sendKeys(getdata(1, 1));

		WebElement login = driver.findElement(By.id("login"));
		submit(login);

		WebElement location = driver.findElement(By.id("location"));
		select(location, 5);

		WebElement hotels = driver.findElement(By.id("hotels"));
		select(hotels, 3);

		WebElement room = driver.findElement(By.id("room_type"));
		select(room, 3);

		WebElement roomnos = driver.findElement(By.id("room_nos"));
		select(roomnos, 2);

		WebElement checkin = driver.findElement(By.id("datepick_in"));
		checkin.clear();
		checkin.sendKeys(getdata(1, 2));

		WebElement checkout = driver.findElement(By.id("datepick_out"));
		checkout.clear();
		checkout.sendKeys(getdata(1, 3));

		WebElement roomtype = driver.findElement(By.id("adult_room"));
		select(roomtype, 2);

		WebElement childroom = driver.findElement(By.id("child_room"));
		select(childroom, 2);

		WebElement submit = driver.findElement(By.id("Submit"));
		submit(submit);

		WebElement radio = driver.findElement(By.id("radiobutton_0"));
		submit(radio);

		WebElement cont = driver.findElement(By.id("continue"));
		submit(cont);
		
		WebElement firstname = driver.findElement(By.id("first_name"));
		firstname.sendKeys(getdata(1,4));
		
		WebElement lastname = driver.findElement(By.id("last_name"));
		lastname.sendKeys(getdata(1,5));
		
		WebElement address = driver.findElement(By.id("address"));
		address.sendKeys(getdata(1,6));
		
		WebElement cardno = driver.findElement(By.id("cc_num"));
		cardno.sendKeys(getdata(1,7));
		
		WebElement cardtype = driver.findElement(By.id("cc_type"));
		select(cardtype, 2);
		
		WebElement month = driver.findElement(By.id("cc_exp_month"));
		select(month, 9);
		
		WebElement year = driver.findElement(By.id("cc_exp_year"));
		select(year, 11);
		
		WebElement cvv = driver.findElement(By.id("cc_cvv"));
		cvv.sendKeys(getdata(1,8));
		
		WebElement book = driver.findElement(By.id("book_now"));
		submit(book);
		
		
		driver.manage().timeouts().implicitlyWait(5000, TimeUnit.SECONDS);
		WebElement orderno = driver.findElement(By.xpath("//input[@id='order_no']"));
		String text=orderno.getAttribute("value");
		
		System.out.println(text);
		
		File loc1 = new File("C:\\Users\\M\\eclipse-1\\MavenClass\\datas\\Book1.xls");

		FileInputStream stream1 = new FileInputStream(loc1);

		Workbook w1 = new HSSFWorkbook(stream1);

		Sheet s1 = w1.getSheet("Data");

		Row r1 = s1.getRow(1);

		Cell c1 = r1.createCell(9);
				
		c1.setCellValue(text);
		

		FileOutputStream stream2=new FileOutputStream(loc1);
		
		w1.write(stream2);
		
		
		WebElement logout = driver.findElement(By.id("logout"));
		logout.click();
		
		
		driver.quit();
		

	}

}
