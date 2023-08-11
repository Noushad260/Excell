package ExcelUtility;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class AmazonNamePrice {
	public static void main(String[] args) throws EncryptedDocumentException, IOException, InterruptedException {
		// System.setProperty("webdriver.chrome.driver", "F:\\Selenium F
		// Folder\\Selenium_\\Set Chrome\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		driver.get("https://www.amazon.in");
		driver.findElement(By.id("twotabsearchtextbox")).sendKeys("i phone", Keys.ENTER);
		List<WebElement> name = driver
				.findElements(By.xpath("//span[@class=\"a-size-medium a-color-base a-text-normal\"]"));
		List<WebElement> price = driver.findElements(By.xpath("//span[@class=\"a-price-whole\"]"));
		FileInputStream fis = new FileInputStream("./src/Eclips.xlsx");
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sh = wb.getSheet("Sheet5");
		ArrayList a1 = new ArrayList<>();
		ArrayList a2 = new ArrayList<>();
		for (WebElement e : name) {
			a1.add(e.getText());
		}
		for (WebElement e2 : price) {
			a2.add(e2.getText());
		}
		Iterator itr1 = a1.iterator();
		Iterator itr2 = a2.iterator();
		int i = 0;
		while (itr1.hasNext() && itr2.hasNext()) {

			Object obj = itr1.next();
			Object obj1 = itr2.next();
			String s1 = obj1.toString();
			String s2 = obj.toString();
//			System.out.println(s1 + " " + s2);
			Row r1 = sh.createRow(i);
			i++;
			Cell c1 = r1.createCell(1);
			Cell c2 = r1.createCell(0);
			c1.setCellValue(s1);
			c2.setCellValue(s2);
		}

		FileOutputStream out = new FileOutputStream("./src/Eclips.xlsx");
		wb.write(out);
		wb.close();

	}
}
