package utils;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ExcelDataProvider {

	@Test(dataProvider = "test1data")
	public void test1(String Username, String Password, String Status, String Phonenumber, String zipcode) {
		System.out.println(Username + "|" + Password + "|" + Status + "|" + Phonenumber + "|" + zipcode);
	}

	@DataProvider(name = "test1data")
	public Object[][] getdata() {
		String excelpath = System.getProperty("user.dir");
		// testData(excelpath, "sheet1");
		Object data[][] = testData(excelpath + "/Excel/TestData.xlsx", "sheet1");
		return data;
	}

	public Object[][] testData(String excelPath, String sheetName) {
		ExcelUtils excel = new ExcelUtils(excelPath, sheetName);
		int rowCount = excel.getRowCount();
		int colCount = excel.getColCount();

		Object data[][] = new Object[rowCount - 1][colCount];

		for (int i = 1; i < rowCount; i++) {
			for (int j = 0; j < colCount; j++) {
				{
					String CellData = excel.getCellData(i, j);
					// System.out.print(CellData + "|");
					// CellData=(String) data[i - 1][j];
					data[i - 1][j] = CellData;

				}

			}
			// System.out.println();

		}
		return data;

	}

}
