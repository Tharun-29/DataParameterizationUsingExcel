package ExcelAutomation;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class dataExtracter {
	
	DataFormatter formatter = new DataFormatter();

	@Test(dataProvider = "TestDataDrive")
	public void testcasedata(String message1, String message2, String id) {
		System.out.println(message1 + message2 + id);
	}

	@DataProvider(name = "TestDataDrive")
	public Object[][] getData() throws IOException {
		// Object[][] data =
		// {{"Hi","Hello",1},{"How","are",2},{"Be","Confident",3},{"Stay","Positive",4}};

		FileInputStream input = new FileInputStream("C:\\Users\\91807\\OneDrive\\Desktop\\exceldriving.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(input);

		XSSFSheet sheet = workbook.getSheetAt(0);

		int rowCount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int colCount = row.getLastCellNum();

		Object[][] Data = new Object[rowCount - 1][colCount];

		for (int i = 0; i < rowCount - 1; i++) {
			row = sheet.getRow(i + 1);
			for (int j = 0; j < colCount; j++) {
				 XSSFCell cell = row.getCell(j);
				 
				 Data[i][j] = formatter.formatCellValue(cell);
			}
		}

		return Data;
	}

}
