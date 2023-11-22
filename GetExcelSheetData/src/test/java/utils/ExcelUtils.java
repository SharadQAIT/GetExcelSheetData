package utils;

import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	static String Projectpath;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;

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
	public int getRowCount() {
		int rowCount = 0;
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
	public int getColCount() {
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
	public String getCellData(int rowNum, int colNum) {
		String celldata = null;
		try {
			Projectpath = System.getProperty("user.dir");
			workbook = new XSSFWorkbook(Projectpath + "/Excel/TestData.xlsx");
			sheet = workbook.getSheet("sheet1");
			// converting data from one format to another, such as formatting dates,
			// numbers, or strings.
			DataFormatter formatter = new DataFormatter();
			try {
				celldata = formatter.formatCellValue(sheet.getRow(rowNum).getCell(colNum));
			} catch (Exception e) {
				e.printStackTrace();
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			e.getCause();
			e.getMessage();
		}
		// get celldata value
		return celldata;

	}

}
