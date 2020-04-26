package dataDriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Data {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
FileInputStream fil = new FileInputStream("C:\\Users\\User\\Desktop\\test\\Selenium-Java\\dataDrivenExcel.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook(fil);
		
		int sheets = workbook.getNumberOfSheets();
		for(int i=0; i<sheets; i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
				XSSFSheet sheet = workbook.getSheetAt(i); 	// gets collection rows
				
				// Identify Testcases column by scanning entire first row
				Iterator<Row> rows = sheet.iterator();		//has all the rows
				Row firstrow =rows.next();	// will move to first row and has all the cells in the row. 2nd rows.next() moves to 2nd row
				Iterator<Cell> cells = firstrow.cellIterator();  	//has all the cells. so cells.next() has ist cell
				//System.out.println(cells.next().getStringCellValue());				// has 1st cell*/
				while(cells.hasNext()) {
					System.out.println(cells.next().getStringCellValue());
				}
	}
}
	}
}
