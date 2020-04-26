package dataDriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
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
	
	
	public ArrayList<String> getData(String testCaseName) throws IOException
	{
		ArrayList<String> a = new ArrayList<String>();
		FileInputStream fil = new FileInputStream("C:\\Users\\User\\Desktop\\test\\Selenium-Java\\dataDrivenExcel.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook(fil);
		
		int sheets = workbook.getNumberOfSheets();
		for(int i=0; i<sheets; i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
				XSSFSheet sheet = workbook.getSheetAt(i); 	// gets collection rows
				
				// Identify Testcases column by scanning entire first row
				Iterator<Row> rows = sheet.iterator();		//has all the rows
				Row firstrow = rows.next();	// will move to first row and has all the cells in the row. 2nd rows.next() moves to 2nd row
				Iterator<Cell> cells = firstrow.cellIterator();  	//has all the cells. so cells.next() has ist cell
				//Cell firstcel = cells.next();				// has 1st cell
				
				// To check if next cell is present, we use while loop + hasnext()
				int k=0;
				int column = 0;
				while(cells.hasNext()) {			// this only check if there is next cell present to the right, it does not move to it
					Cell cell = cells.next();		//this while loop actually move to the cells if present
					System.out.println(cell.getStringCellValue());
					if(cell.getStringCellValue().equalsIgnoreCase("Testcases")) {		 
						
						// cell identified
						column=k;						//here we have ist column row index
						
					}
					k++;
					
				}
				System.out.println(column);
				
				//Once column is identified, then scan the entire Testcases column to identify purchase Testcases row
				//Iterator<Row> rows = sheet.iterator();
				while(rows.hasNext()) {
					Row row = rows.next();			// gets the rows downward
					if(row.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName)) {
						
						// After you grab Purchase row, pull the data of that row and feed into test case
						Iterator<Cell> purchaseCells = row.cellIterator();
						while(purchaseCells.hasNext()) {
							//System.out.println(purchaseCells.next().getStringCellValue());
	//Instead of printing, we can feed the data into our test cases using Array List
	// Create Array list at the top and send the data to the array
	a.add(purchaseCells.next().getStringCellValue()); 		// it is now properly stored in the array list
	
	Cell c = purchaseCells.next();
	if(c.getCellType()==CellType.STRING)
	{
		a.add(c.getStringCellValue());
	}
	else 
	{
		
		a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
	}
	// We can actually wrap this whole thing in a Method
							
						}
					}
					
				}
			}
		}
		return a;
	}
	
	
//   Now we can call this method from another class (test sample) where we can use the data, either as a sendkeys username and password to a web site.

}
