package testscripts;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Test4 {
	public static void main(String[] args) throws Exception {
		String xlPath = "./xldata/Test Data.xlsx";
		
	
	FileInputStream fis=new FileInputStream(xlPath);
	Workbook wb=WorkbookFactory.create(fis);

	Sheet s=wb.getSheet("Sheet3");
	 int rowCount = s.getLastRowNum(); // index of last row
	 System.out.println(rowCount);
	 int cellCountRow1=s.getRow(0).getLastCellNum(); //no of cells
	 System.out.println(cellCountRow1);
	 
	 
}
}

