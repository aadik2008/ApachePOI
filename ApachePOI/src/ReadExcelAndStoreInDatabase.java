

import java.io.*;
import java.sql.*;
import java.util.*;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
 

public class ReadExcelAndStoreInDatabase {
 
    public static void main(String[] args) {
    	
    	
    	try {
			// File file = new File("C:\\employee.xlsx"); // creating a new file instance
			File file = new File(".\\datafiles\\Emp.xlsx"); // creating

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
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

        int batchSize = 20;
 
        try {
                         
            FileInputStream inputStream = new FileInputStream(".\\datafiles\\Emp.xlsx");
 
            Workbook workbook = new XSSFWorkbook(inputStream);
 
            Sheet firstSheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = firstSheet.iterator();
 
            Connection connection = DriverManager.getConnection("jdbc:mysql://localhost:3306/employee", "root", "@dity@ranu123");
            connection.setAutoCommit(false);
  
            String sql = "INSERT INTO emp (Serial_No,Employee_Name,Designation) VALUES (?, ?, ?)";
            PreparedStatement statement = connection.prepareStatement(sql);    
             
            int count = 0;
             
            rowIterator.next(); // skip the header row
             
            while (rowIterator.hasNext()) {
                Row nextRow = rowIterator.next();
                Iterator<Cell> cellIterator = nextRow.cellIterator();
 
                while (cellIterator.hasNext()) {
                    Cell nextCell = cellIterator.next();
 
                    int columnIndex = nextCell.getColumnIndex();
 
                    switch (columnIndex) {
                    case 0:
                    	int Serial_No= (int) nextCell.getNumericCellValue();
                    	statement.setInt(1, Serial_No);
                    	break;
                    case 1:
                        String Employee_Name = nextCell.getStringCellValue();
                        statement.setString(2, Employee_Name);
                       
                    case 2:
                        String Designation = nextCell.getStringCellValue();
                        statement.setString(3, Designation);
                    }
 
                }
                 
                statement.addBatch();
                 
                if (count % batchSize == 0) {
                    statement.executeBatch();
                }           
 
            }
 
            workbook.close();
             
            /*For Proper Understanding:---it will Submit a batch of commands to the database for execution and 
             * if all commands execute successfully, returns an array of update counts.*/
            statement.executeBatch();
  
            connection.commit();
            connection.close();
             
            
            System.out.printf("Done Successfully!!!");
             
        } catch (IOException e1) {
            System.out.println("Error in file");
            e1.printStackTrace();
        } catch (SQLException e2) {
            System.out.println("Database error");
            e2.printStackTrace();
        }
 
    }
}


 