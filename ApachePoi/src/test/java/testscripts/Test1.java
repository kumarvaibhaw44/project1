package testscripts;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Test1 {
	public static void main(String[] args) throws Exception {
		String xlPath = "./xldata/Test Data.xlsx";
		
	
	FileInputStream fis=new FileInputStream(xlPath);
	Workbook wb=WorkbookFactory.create(fis);
	Sheet s = wb.getSheet("Sheet1");
	 Row r=s.getRow(0);
	 Cell c=r.getCell(0);
	 String v = c.toString();
	// wb.getSheet("Sheet1").getRow(0).getCell(0).toString()
	 System.out.println(v);
	
	
	
	
	
	}

}
