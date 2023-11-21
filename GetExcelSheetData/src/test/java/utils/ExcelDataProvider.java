package utils;

public class ExcelDataProvider {

	public static void main(String[] args) {
		String excelpath = System.getProperty("user.dir");

		testData(excelpath + "/Excel/TestData.xlsx", "sheet1");
	}

	@SuppressWarnings("static-access")
	public static void testData(String excelPath, String sheetName) {
		ExcelUtils excel = new ExcelUtils(excelPath, sheetName);
		int rowCount = excel.getRowCount();
		int colCount = excel.getColCount();

		System.out.println("rowcount: " + rowCount);
		System.out.println("colCount: " + colCount);

		for (int i = 1; i < rowCount; i++) {
			for (int j = 0; j < colCount; j++) {
				String CellData1 = excel.getCellDataString(i, j);
				int CellData2 = excel.getCellDataNumeric(i, j);
				System.out.print(CellData1 + "|" + CellData2);
			}
			System.out.println();

		}

	}

}
