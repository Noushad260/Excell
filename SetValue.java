package ExcelUtility;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class SetValue {
public static void main(String[] args) throws EncryptedDocumentException, IOException {
	
	
	FileInputStream fis = new FileInputStream("./src/Excel File.xlsx");
	Workbook wb = WorkbookFactory.create(fis);
	Sheet sh=wb.getSheet("Sheet1");
	
	Row r2=sh.createRow(7);
	Cell c2=r2.createCell(2);
	c2.setCellValue("MD noushad ansri");
	FileOutputStream out=new FileOutputStream("./src/Excel File.xlsx");
	wb.write(out);
	wb.close();
	
} 

}
