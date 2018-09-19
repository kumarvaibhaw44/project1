package library;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelLib {
public static String readExcelData(String path,String sheet,int row,int cell) throws Exception {
	String v = "";


	try {
		FileInputStream fis=new FileInputStream(path);
		Workbook wb=WorkbookFactory.create(fis);
		 v=wb.getSheet("Sheet").getRow(row).getCell(cell).toString();
	} catch (Exception e) {
		
	}
	return v;
}
public static void writeExcelData(String path,String sheet,int row,int cell,String CellData) {

	try {
		FileInputStream fis=new FileInputStream(path);
		Workbook wb=WorkbookFactory.create(fis);
		 wb.getSheet("Sheet").createRow(row).createCell(cell).setCellValue(CellData);
		 FileOutputStream fos=new FileOutputStream(path);
		 wb.write(fos);
	} catch (Exception e) {
	
	}
	 
}

public static int getRowCount(String path,String sheet) {
	int rowCount=0;
	

	try {
		FileInputStream fis=new FileInputStream(path);
		Workbook wb=WorkbookFactory.create(fis);
		rowCount= wb.getSheet("Sheet").getLastRowNum();
	} catch (Exception e) {
		
	
}
	return rowCount;

}
public static int getCellCount(String path,String sheet,int row) {
	int cellCount=0;
	

	try {
		FileInputStream fis=new FileInputStream(path);
		Workbook wb=WorkbookFactory.create(fis);
	cellCount= wb.getSheet("Sheet").getRow(row).getLastCellNum();
	} catch (Exception e) {
		
	
}
	return cellCount;
}}
