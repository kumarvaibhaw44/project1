package testscripts;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Test5 {
	public static void main(String[] args) throws Exception {
		String xlPath = "./xldata/Test Data.xlsx";
		
	
	FileInputStream fis=new FileInputStream(xlPath);
	Workbook wb=WorkbookFactory.create(fis);

	Sheet s=wb.getSheet("Sheet3");
	 int rowCount = s.getLastRowNum();
	 for(int i=0;i<=rowCount;i++) {
		 int cellCount=s.getRow(i).getLastCellNum();

		 for(int j=0;j<cellCount;j++) {
			 System.out.print(s.getRow(i).getCell(j)+ "  ");
		 }
		 System.out.println();
	 }
}
}