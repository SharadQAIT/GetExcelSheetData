package utils;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	static String Projectpath;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	public int celldata2;
	
	public ExcelUtils(String excelpath, String sheetName) {
		try {
			Projectpath = System.getProperty("user.dir");
			workbook = new XSSFWorkbook(excelpath);
			sheet = workbook.getSheet(sheetName);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/*
	 * This is row count code
	 * 
	 */
	public static int getRowCount() {
		int rowCount=0;
		try {
			rowCount = sheet.getPhysicalNumberOfRows();
			System.out.println("No of rows: " + rowCount);
		} catch (Exception e) {
			e.getCause();
			e.printStackTrace();

		}
		return rowCount;

	}

	/*
	 * This is column count code
	 * 
	 */
	
	public static int getColCount() {
		int colCount = 0;
		try {
			colCount = sheet.getRow(0).getPhysicalNumberOfCells();
			System.out.println("No of cols: " + colCount);
		} catch (Exception e) {
			e.printStackTrace();
			e.getCause();
		}
		return colCount;

		// return colCount;
	}

			/*
			 * Get string data from excel file
			 * 
			 */
	public static String getCellDataString(int rowNum, int colNum) {
		String cellData1 =null;
		try {
			Projectpath = System.getProperty("user.dir");
			workbook = new XSSFWorkbook(Projectpath + "/Excel/TestData.xlsx");
			sheet = workbook.getSheet("sheet1");
			cellData1 = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
			//System.out.println("1st coulmn of cell data: " + cellData1);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			e.getCause();
			e.getMessage();
		}
		return cellData1;
	}
	
	/*
	 * Get interger data from excel file
	 */

	public static double getCellDataNumeric(int rowNum, int colNum) {
		double celldata2= 0;
		try {
			Projectpath = System.getProperty("user.dir");
			workbook = new XSSFWorkbook(Projectpath + "/Excel/TestData.xlsx");
			sheet = workbook.getSheet("sheet1");
			celldata2 = sheet.getRow(rowNum).getCell(rowNum).getNumericCellValue();
			//System.out.println("|" + cellData2);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			e.getCause();
			e.getMessage();
		}
		return celldata2;

	}
}
