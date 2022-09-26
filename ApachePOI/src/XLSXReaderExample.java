/*
 
5.) Using Apache poi library:-

- Read an excel sheet.
- Add proper validation and handle all the errors scenarios. For example: file not found, invalid format etc.
- Insert valid records into in -memory/mysql/any other Database

*/

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class XLSXReaderExample {
	public static void main(String[] args) {
		try {
			//File file = new File("C:\\employee.xlsx"); // creating a new file instance
			File file = new File(".\\datafiles\\Countries.xlsx"); // creating
			
			FileInputStream files = new FileInputStream(file); 
			XSSFWorkbook workbook = new XSSFWorkbook(files);
			XSSFSheet sheet = workbook.getSheetAt(0); 
			Iterator<Row> itr = sheet.iterator(); 
			while (itr.hasNext()) {
				Row row = itr.next();
				Iterator<Cell> cellIterator = row.cellIterator(); 
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					switch (cell.getCellType()) {
					case STRING: 
						System.out.print(cell.getStringCellValue() + "\t\t\t");
						break;
					case NUMERIC: 
						System.out.print(cell.getNumericCellValue() + "\t\t\t");
						break;
					default:
					}
				}
				System.out.println("");
			}
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
}
