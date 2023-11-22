package utils;

public class ExcelDataProvider extends ExcelUtils{

	public ExcelDataProvider(String excelpath, String sheetName) {
		super(excelpath, sheetName);
		// TODO Auto-generated constructor stub
	}

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
				if (j < i) 
				{
					String CellData1 = excel.getCellDataString(i, j);
					System.out.print(CellData1 + "|");
				} 
				else 
				{
					double CellData2 = excel.getCellDataNumeric(i, j);
					System.out.print(CellData2);
				}

			}
			System.out.println();

		}

	}

}
