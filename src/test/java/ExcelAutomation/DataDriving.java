package ExcelAutomation;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriving {

	public ArrayList<String> getData(String testcaseName) throws IOException {
		ArrayList<String> list = new ArrayList<String>();

		FileInputStream input = new FileInputStream("C:\\Users\\91807\\OneDrive\\Desktop\\data.xlsx");
		
		/*
		 * Steps - 
		 * 
		 * 1. Create Object for XSFF WorkBook class
		 * 2. Get Access to Sheet
		 * 3. Get Access to all rows of Sheet
		 * 4. Access to Specific row from all rows
		 * 5. Get Access to all cells of Row
		 * 6. Access the data from Excel into Array
		 */
		
		
		//1. Create Object for XSFF WorkBook class
		XSSFWorkbook workbook = new XSSFWorkbook(input);

		int numberOfSheets = workbook.getNumberOfSheets();

		for (int i = 0; i < numberOfSheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("sheet1")) {
				
				//2. Get Access to Sheet
				XSSFSheet sheet = workbook.getSheetAt(i);
                 
				//3. Get Access to all rows of Sheet
				Iterator<Row> rows = sheet.iterator();
				Row firstRow = rows.next();

				Iterator<Cell> columns = firstRow.cellIterator();

				int k = 0;
				int columnIndex = -1;

				while (columns.hasNext()) {
					Cell column = columns.next();
					if (column.getStringCellValue().equalsIgnoreCase(testcaseName)) {
						columnIndex = k;
						break;
					}
					k++;
				}

				if (columnIndex == -1) {
					workbook.close();
					throw new IllegalArgumentException("Test case column not found");
				}

				while (rows.hasNext()) {
					Row row = rows.next();
					if (row.getCell(columnIndex).getStringCellValue().equalsIgnoreCase("fridge")) {
						Iterator<Cell> cellIterator = row.cellIterator();

						while (cellIterator.hasNext()) {
							Cell columnTraverse = cellIterator.next();

							if(columnTraverse.getCellType() == org.apache.poi.ss.usermodel.CellType.STRING) {
								list.add(columnTraverse.getStringCellValue());
							} else {
								list.add(NumberToTextConverter.toText(columnTraverse.getNumericCellValue()));
							}

						}
					}
				}
			}
		}
		workbook.close();
		return list;
	}
}
