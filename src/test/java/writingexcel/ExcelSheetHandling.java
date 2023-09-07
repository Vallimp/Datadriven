package writingexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelSheetHandling {
	
	public static void writeExcelSheet() throws IOException {
		//workbook
	XSSFWorkbook workbook = new XSSFWorkbook();
	//create sheet inside the workbook
	XSSFSheet worksheet = workbook.createSheet("Sheet1");
	
	int rowNum = 0;
	//rows
	for (int r=0; r<=10; r++) {
		//Import Row
		Row row = worksheet.createRow(rowNum++);
		int colNum = 0;
		
		//columns
		for (int c=0; c<=10; c++) {
		//Import Cell -apache poi
		Cell cell = row.createCell(colNum++);
		//setCellValue...select string as we are going to enter strings
		cell.setCellValue("Row " +r + "Column" +c);
		}
	}
	//user.dir gives current project directory plus the exceldata file path
	String path = System.getProperty("user.dir")+"/src/test/resources/TestData/demo.xlsx";
	File ExcelFile = new File(path);
	//fos has to be accessible by finlly class so declare it outside try catch
	FileOutputStream fos = null;
    //Telling where to write the data	
	try {
		 fos = new FileOutputStream(ExcelFile);
		 //write into workbook
		 workbook.write(fos);
		 //after writing close the workbook
		 workbook.close();
	} catch (FileNotFoundException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
	//finally block
	finally {
		fos.close();
	}
}
	public static void readExcelSheet() throws IOException {
		//to read from an excel file, first specify the file path
		String path = System.getProperty("user.dir")+"/src/test/resources/TestData/demo.xlsx";
		File ExcelFile = new File(path);
		//to read from an excel file we need a file input stream
		
			FileInputStream fis = new FileInputStream(ExcelFile);
			//here we are not creating the workbook but use the existing input steam to read the excel
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		//get the sheet from the workbook on the right side, using the same sheet from above method
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		//import iterator from java.util, it will automatically find the number of rows and columns
		Iterator<Row> row = sheet.rowIterator();
		//use while loop so you dont hav to worry about the number of rows
		//whenever there i a next row, it has to iterate
		while(row.hasNext()){
			Row currentrow = row.next();
			//do the same fr cell also
			Iterator<Cell> cell = currentrow.cellIterator();
		while(cell.hasNext()) {
			Cell currentcell = cell.next();
			//get string value of each cell
			System.out.print(currentcell.getStringCellValue()+ " ~ ");
			}
		//for space between rows
		System.out.println();
		}
		//to simultaneously write in the same excel sheet say 12th row and 13th column
		Row newrow = sheet.createRow(12);
		Cell newcell = newrow.createCell(13);
		//it has to write Valli in 12th row 13th column in the demo.xlsx
		newcell.setCellValue("Valli");
		//create output stream
		FileOutputStream fos = new FileOutputStream(ExcelFile);
		workbook.write(fos);
		workbook.close();
	}
	
	public static void main(String[] args) throws IOException {
		writeExcelSheet();
		readExcelSheet();
	}
}
	

