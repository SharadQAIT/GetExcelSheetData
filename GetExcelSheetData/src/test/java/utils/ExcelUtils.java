package utils;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	static String Projectpath;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	
	public ExcelUtils(String excelpath, String sheetName)
	{
		try {
			Projectpath = System.getProperty("user.dir");
			workbook = new XSSFWorkbook(excelpath);
			sheet= workbook.getSheet(sheetName); 
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public static void main(String[] args) {
		getRowCount();
		getCellDataString(1,0);
		getCellDataNumeric(1,1);


	}
	public static void getRowCount()
	{
	int rowCount=sheet.getPhysicalNumberOfRows();
	System.out.println("No of rows: "+rowCount);
}
	public static void getCellDataString(int rowNum, int colNum )
	{
	try {
		Projectpath = System.getProperty("user.dir");
		workbook = new XSSFWorkbook(Projectpath+"/Excel/TestData.xlsx");
		sheet= workbook.getSheet("sheet1"); 
		String cellData1=sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
		System.out.println("1st coulmn of cell data: "+cellData1);
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
		e.getCause();
		e.getMessage();
	}
	}
	public static void getCellDataNumeric(int rowNum, int colNum )
	{
	try {
		Projectpath = System.getProperty("user.dir");
		workbook = new XSSFWorkbook(Projectpath+"/Excel/TestData.xlsx");
		sheet= workbook.getSheet("sheet1"); 
		Double cellData2=sheet.getRow(rowNum).getCell(rowNum).getNumericCellValue();
		System.out.println("2nd coulmn of cell data: "+cellData2);
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
		e.getCause();
		e.getMessage();
	}

	}
}
