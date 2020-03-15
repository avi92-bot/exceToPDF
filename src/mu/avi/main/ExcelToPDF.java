package mu.avi.main;

import java.io.FileInputStream;
import java.util.Date;
import java.util.Objects;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToPDF {
	
	public static final String XLS_EXTENSION = "xls";
	public static final String XLSX_EXTENSION = "xlsx";
	
	public static void main(String[] args) {
		try {
			Objects.requireNonNull(args);
	        int count = 0;
	        if(args.length == 0) {
	        	System.out.println("Wrong argument passed");
	        }else {
				for (int i = 0; i < args.length; i++) {
		            switch (args[i]) {
		            case "-p": //set password 
		                i++;
		                System.out.println("Password is: " + args[i]);
		                break;
		            default:
		            	boolean isXLS = checkExcelFormat(args[i]);
		    			
		    			if(isXLS) {
		    				System.out.println("File extension is: "+ XLS_EXTENSION);
		    				readFromXLSExcel(args[i]);
		    			}else {
		    				System.out.println("File extension is: "+ XLSX_EXTENSION);
		    				readFromXLSXExcel(args[i]);
		    			}
		    			count++;
		    			
		    			System.out.println("file number: "+ count + "finished");
		            }
				}
	        }

		}catch (Exception e) {
			e.printStackTrace();
			System.out.println("Wrong argument passed");
		}
	}
	
	public static void readFromXLSExcel(String file) { 
		try {
			HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
			int numberOfSheet = myExcelBook.getNumberOfSheets();
			for (int i = 0; i < numberOfSheet; i++) {
				HSSFSheet sheet = myExcelBook.getSheetAt(i);
				System.out.println("Sheet number: " + i + " and name: " + sheet.getSheetName());
				int lastRowNumber = sheet.getLastRowNum()+1;
				for (int j = 0; j < lastRowNumber; j++) {
					Row row = sheet.getRow(j);
				    if(row == null) {
				        continue;
				    }
				    int lastCellNumber = row.getLastCellNum()+1;
				    for (int k = 0; k < lastCellNumber; k++) {
				    	Cell cell = row.getCell(k);
				        if(cell == null) {
				            continue;
				        }
				        
				        if(cell.getCellType() == HSSFCell.CELL_TYPE_STRING){ 
							String name = cell.getStringCellValue(); 
							System.out.println("NAME : " + name); 
						} 
						if(cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC){ 
							Date birthdate = cell.getDateCellValue(); 
							System.out.println("DOB :" + birthdate); 
						}
				    }
				}
			}
			myExcelBook.close();
		}catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public static void readFromXLSXExcel(String file){ 
		try {
			XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(file)); 
			int numberOfSheet = myExcelBook.getNumberOfSheets();
			for (int i = 0; i < numberOfSheet; i++) {
				XSSFSheet sheet = myExcelBook.getSheetAt(i);
				System.out.println("Sheet number: " + i + " and name: " + sheet.getSheetName());
				int lastRowNumber = sheet.getLastRowNum()+1;
				for (int j = 0; j < lastRowNumber; j++) {
				    Row row = sheet.getRow(j);
				    if(row == null) {
				        continue;
				    }
				    int lastCellNumber = row.getLastCellNum()+1;
				    for (int k = 0; k < lastCellNumber; k++) {
				        Cell cell = row.getCell(k);
				        if(cell == null) {
				            continue;
				        }
				        
				        if(cell.getCellType() == HSSFCell.CELL_TYPE_STRING){ 
							String name = cell.getStringCellValue(); 
							System.out.println("NAME : " + name); 
						} 
						if(cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC){ 
							Date birthdate = cell.getDateCellValue(); 
							System.out.println("DOB :" + birthdate); 
						}
				    }
				}
			}
			myExcelBook.close(); 
		}catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public static boolean checkExcelFormat(String fileName) {
		
		String excelName = null;
		int lastIndexOfDot = 0;
		int lastIndexOfSlash = 0;
		int fileNameLength = 0;
		
		try {
			Objects.requireNonNull(fileName);
			lastIndexOfDot = fileName.lastIndexOf('.');
			lastIndexOfSlash = fileName.lastIndexOf('/');
			fileNameLength = fileName.length();
	        
	        if(lastIndexOfSlash > 0) {
	        	excelName = fileName.substring(lastIndexOfSlash + 1, fileNameLength);
	        	System.out.println("Excel to be process: " + excelName);
	        }else {
	        	System.out.println("Excel to be process: " + fileName);
	        }

	        String fileExtension = fileName.substring(lastIndexOfDot+1, fileNameLength);
	        if(fileExtension.equalsIgnoreCase(XLS_EXTENSION)) {
	        	return true;
	        }
	        
		}catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		
		return false;
	}

	

}
