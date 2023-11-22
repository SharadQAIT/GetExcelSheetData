package utils;

public class ExcelUtilsDemo {

	@SuppressWarnings("static-access")
	public static void main(String[] args) {
		
		String projectpath = System.getProperty("user.dir");
		ExcelUtils excel=new ExcelUtils(projectpath+"/Excel/TestData.xlsx", "sheet1");
		excel.getRowCount();
		excel.getColCount();
	//	excel.getCellDataString(1, 0);
	//	excel.getCellDataNumeric(1, 1);
	}
}
