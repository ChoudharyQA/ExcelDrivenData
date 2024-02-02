import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {
	
	// First we need access to the excel sheet
	// Then we need to go to the test data sheet
	// Then in test data sheet, we need to check all the columns in which we have purchase 
	// Once we reached to the desired column, then we need to jump to the purchase row
	// And once we reached to the purchase row then we need to feed purchase row data into our test cases.
	
	public ArrayList<String> getData(String testCaseName) throws IOException {
ArrayList <String> a = new ArrayList<String>();
		
		FileInputStream fis = new FileInputStream("C://Users//admin//Desktop//Excel.xlsx");
		XSSFWorkbook workBook = new XSSFWorkbook(fis);
		
		int sheets = workBook.getNumberOfSheets();
		
		for(int i=0;i<sheets;i++) {
			
			if(workBook.getSheetName(i).equalsIgnoreCase("testdata")) {
				
				
				XSSFSheet sheet =  workBook.getSheetAt(i);
				// Here we will identify the Testcases column by scanning the complete row 
				
				Iterator <Row> rows = sheet.iterator(); // Sheet is the collection of rows
			    Row	firstRow = rows.next();
				Iterator <Cell> ce=firstRow.cellIterator(); // Row is the collection of cells
				int k=0;
				int column = 0;
				while(ce.hasNext()) {
					
					Cell value = ce.next();
					if(value.getStringCellValue().equalsIgnoreCase("TestCases")) {
						
						column =k;
					}
					
					k++;
				}
				
				System.out.println(column);
				
				// Once column identified (TestCases), we need to scan the complete column to find "Purchase" row
				
				while(rows.hasNext()) {
					
					Row r =rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName)) {
						
						Iterator <Cell> cv=r.cellIterator();
						while(cv.hasNext()) {
							
							Cell c=cv.next();
							if(c.getCellType()==CellType.STRING) {
								
								
								a.add(c.getStringCellValue());
							}
							else {
								
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
						}
						
						
					}
					
					
					
					
					
					
				}
								
				
			}
			
		}
		
		return a;
		
		}

	public static void main(String[] args) throws IOException {
		

	}

}
