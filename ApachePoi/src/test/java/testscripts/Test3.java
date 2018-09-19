package testscripts;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Test3 {
	public static void main(String[] args) throws Exception {
		String xlPath = "./xldata/Test Data.xlsx";
		
	
	FileInputStream fis=new FileInputStream(xlPath);
	Workbook wb=WorkbookFactory.create(fis);

	Sheet s=wb.getSheet("Sheet3");
	
	 Cell c;
	 for(int i=0;i<3;i++) {
		 for(int j=0;j<3;j++) {
			 c=s.getRow(i).getCell(j);
			 System.out.print(c+"  ");
			 
		 }
		 System.out.println();
	 }
	 
	}
}
