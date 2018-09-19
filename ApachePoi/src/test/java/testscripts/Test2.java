package testscripts;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Test2 {
	public static void main(String[] args) throws Exception {
		String xlPath = "./xldata/Test Data.xlsx";
		
	
	FileInputStream fis=new FileInputStream(xlPath);
	Workbook wb=WorkbookFactory.create(fis);

	Sheet s=wb.createSheet("Sheet2");
	 Row r=s.createRow(0);
	 Cell c=r.createCell(0);
	 c.setCellValue("Selenium");
	 FileOutputStream fos=new FileOutputStream(xlPath);
	 System.out.println(fos);
	 wb.write(fos);
	 
	
	}
}
