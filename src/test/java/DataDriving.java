import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriving {
     
	
	//Test Case -
	//1.Identify TestCases Column by scanning the entire first row
	//2.Once column is identified then scan the entire Test case column to identify purchase test case
	//3.After we Grab Purchase Test case row -> we pull all the data of that row and feed into test
	
	
	public static void main(String[] args) throws IOException {
       
		
		FileInputStream input = new FileInputStream("C:\\Users\\91807\\OneDrive\\Desktop\\data.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(input);
		int numberOfSheets = workbook.getNumberOfSheets();
             
		for(int i=0;i<numberOfSheets;i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase("sheet1")) {
			XSSFSheet sheet = workbook.getSheetAt(i);
			
			
			//Identify TestCases Column by scanning the entire first row
			Iterator<Row> rows = sheet.iterator();
			Row firstrow = rows.next();
			Iterator<Cell> columns = firstrow.cellIterator();
			
			int k=0;
			int col = 0;
			
			while(columns.hasNext()) {
				Cell column = columns.next();
				
        
				
				if(column.getStringCellValue().equalsIgnoreCase("testcase")) {
	                col = k;
	                break;
				}
				k++;
			}
			System.out.println(col);
			
			//Once column is identified then scan the entire Test case column to identify purchase test case
			while(rows.hasNext()) {
				Row r = rows.next();
				if(r.getCell(col).getStringCellValue().equalsIgnoreCase("purchase")) {
			// After we Grab Purchase Test case row -> we pull all the data of that row and feed into test
				       Iterator<Cell> testcell = r.cellIterator();
				   
				       while(testcell.hasNext()) {
				    	   Cell testcolumn = testcell.next();
				    	   System.out.println(testcolumn.getStringCellValue());
				       }
				       
				}
			}
			
			
		}
		
		}
	}

}
