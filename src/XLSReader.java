import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class XLSReader {
	try{
	     
	    FileInputStream file = new FileInputStream(new File("scouting"));
	     
	    //Get the workbook instance for XLS file 
	    HSSFWorkbook workbook = new HSSFWorkbook(file);
	 
	    //Get first sheet from the workbook
	    HSSFSheet sheet = workbook.getSheetAt(0);
	     
	    //Iterate through each rows from first sheet
	    Iterator<Row> rowIterator = sheet.iterator();
	    while(rowIterator.hasNext()) {
	        Row row = rowIterator.next();
	         
	        //For each row, iterate through each columns
	        Iterator<Cell> cellIterator = row.cellIterator();
	        while(cellIterator.hasNext()) {
	             
	            Cell cell = cellIterator.next();
	             
	            switch(cell.getCellType()) {
	                case Cell.CELL_TYPE_BOOLEAN:
	                    System.out.print(cell.getBooleanCellValue() + "\t\t");
	                    break;
	                case Cell.CELL_TYPE_NUMERIC:
	                    System.out.print(cell.getNumericCellValue() + "\t\t");
	                    break;
	                case Cell.CELL_TYPE_STRING:
	                    System.out.print(cell.getStringCellValue() + "\t\t");
	                    break;
	            }
	        }
	        System.out.println("");
	    }
	}
	    file.close();
	    FileOutputStream out = 
	        new FileOutputStream(new File("C:\\test.xls"));
	    workbook.write(out);
	    out.close();
	     
	} catch (FileNotFoundException e1) {
	    e1.printStackTrace();
	}catch (IOException e2) {
	    e1.printStackTrace();
	}
}
