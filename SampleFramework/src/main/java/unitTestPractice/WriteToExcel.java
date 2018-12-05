package unitTestPractice;

import java.io.*;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.text.DecimalFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import unitTestPractice.Practice;


public class WriteToExcel {

	static String filepath = null;
	Object[][] data = null;
	public FileInputStream fis = null;
    public XSSFWorkbook workbook = null;
    public XSSFSheet sheet = null;
    public XSSFRow row = null;
    public XSSFCell cell = null;
	public static void main(String ar[]) throws Exception {
		//WriteToExcel rw = new WriteToExcel(Constants.Path_TestData);
		//WriteToExcel.writeDataToExcel(filepath);
		//rw.readDataFromExcel();
		WriteToExcel.setcelldata("C:\\Users\\Public\\Documents\\SampleFramework\\TestSuites&TestCases\\New folder\\Test.xlsx", "Test Data", "Result1", "Test data pSuccess");
		//WriteToExcel.writecolumn();
		
		
	}
	
	 public String getCellData(String sheetName, String colName, int rowNum)
	    {
		  try
	        {
	        	fis = new FileInputStream("C:\\Users\\Public\\Documents\\SampleFramework\\TestSuites&TestCases\\New folder\\Test.xlsx");
		        workbook = new XSSFWorkbook(fis);
	            int col_Num = -1;
	            sheet = workbook.getSheet(sheetName);
	            row = sheet.getRow(0);
	            for(int i = 0; i < row.getLastCellNum(); i++)
	            {
	                if(row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
	                    col_Num = i;
	            }
	 
	            row = sheet.getRow(rowNum - 1);
	            cell = row.getCell(col_Num);
	            
	 
	            if(cell.getCellTypeEnum() == CellType.STRING)
	                return cell.getStringCellValue();
	            else if(cell.getCellTypeEnum() == CellType.NUMERIC || cell.getCellTypeEnum() == CellType.FORMULA)
	            {
	                String cellValue = String.valueOf(cell.getNumericCellValue());
	                if(HSSFDateUtil.isCellDateFormatted(cell))
	                {
	                    DateFormat df = new SimpleDateFormat("dd/MM/yy");
	                    Date date = cell.getDateCellValue();
	                    cellValue = df.format(date);
	                }
	                return cellValue;
	            }else if(cell.getCellTypeEnum() == CellType.BLANK)
	                return "";
	            else
	                return String.valueOf(cell.getBooleanCellValue());
	        }
	        catch(Exception e)
	        {
	            e.printStackTrace();
	            return "row "+rowNum+" or column "+colName+" does not exist  in Excel";
	        }
	      
	    }
	
	public WriteToExcel() {
		this.filepath = filepath;
	}

	

	public File getFile() throws FileNotFoundException {
		File here = new File(filepath);
		return new File(here.getAbsolutePath());

	}
	
	public static void writecolumn()
	{
		File file =    new File("C:\\Users\\Public\\Documents\\SampleFramework\\TestSuites&TestCases\\New folder\\Test.xlsx");
		 try {
		  // Open the Excel file
        FileInputStream fis = new FileInputStream(file);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
	    XSSFSheet sheet = wb.getSheet("Test Data");
	    int TotalRow;
	    int TotalCol;
	    TotalRow = sheet.getLastRowNum();
	    XSSFRow headerRow = sheet.getRow(0);
	    String result = "";
	    int resultCol = -1;
	    for (Cell cell : headerRow){
	        result = cell.getStringCellValue();
	        if (result.equals("Result"))
	        		{
	            resultCol = cell.getColumnIndex();
	          
	            
	            break;
	         }
	    }
	    if (resultCol == -1){
	        System.out.println("Result column not found in sheet");
	        return;
	    }   
	  //  System.out.println("row " + TotalRow +  " col " +  TotalCol);


	    // Loop through all rows in the sheet
	    // Start at row 1 as row 0 is our header row
	    
	      // }
	    fis.close();
	    FileOutputStream outputStream = new FileOutputStream(file);
	    wb.write(outputStream);
	    outputStream.close();
	}
		 catch (IOException e) {
			    System.out.println("Test data file not found");
			}
	}

	
	
	private static void writeToCell(int rowno, int colno, XSSFSheet sheet, String val) {
		try {
			sheet.getRow(rowno);
			XSSFRow row = sheet.getRow(rowno);
			if (row == null) {
				row = sheet.createRow(rowno);
			}
			XSSFCell cell = row.createCell(colno);
			cell.setCellValue(val);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	
	public static void writeDataToExcel(String file) {
		XSSFWorkbook wb = null;
		XSSFSheet sheet = null;
		FileOutputStream fileOut = null;
		
		
			String excelFileName = file;

			String sheetName = Practice.Sheet_TestData;//name of sheet

			wb = new XSSFWorkbook();
			sheet = wb.createSheet(sheetName);
			DecimalFormat df2 = new DecimalFormat(".##");
			try{

	
			
			writeToCell(0, 0, sheet,  "This is one");// row 1 column 1
			writeToCell(0, 1, sheet,  "this is two");// row 1 column 2
			writeToCell(1, 0, sheet,  "this is three");// row 2 column 1
			writeToCell(1, 1, sheet,  "this is four");// row 2 column 2
			int r = 4;
			
			System.out.println("working fine");
			fileOut = new FileOutputStream(excelFileName);
			wb.write(fileOut);
			
			

			//write this workbook to an Outputstream.
		} catch (Exception e) {
			e.printStackTrace();
		} finally {

			try {
				fileOut.flush();
				fileOut.close();
			} catch (IOException ex) {
				Logger.getLogger(WriteToExcel.class.getName()).log(Level.SEVERE, null, ex);
			}

		}
	}

	public static void setcelldata(String path, String sheetName, String colName, String data) throws Exception {
		// create an object of Workbook and pass the FileInputStream object into
		// it to create a pipeline between the sheet and eclipse.
		FileInputStream fis = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet(sheetName);
		int col_Num = -1;
		XSSFRow row = sheet.getRow(0);
		//row.createCell(0).setCellValue(colName);
		
		for (int i = 0; i < row.getLastCellNum(); i++) {
			/*if (row.getCell(i).getStringCellValue().trim().equals(colName)) {*/
			if (row.getCell(i).getStringCellValue().trim().equals(colName)) {
				col_Num = i;
			}
		}
		// int rowNum =ExcelLibrary.getRowCount(sheetName);
		int rowNum = 1;

		//sheet.autoSizeColumn(col_Num);
		row = sheet.getRow(rowNum);
		if (row == null)
			row = sheet.createRow(rowNum);
		

		XSSFCell cell = row.getCell(col_Num);
		if (cell == null)
			cell = row.createCell(col_Num);
		cell.setCellType(cell.CELL_TYPE_STRING);
		cell.setCellValue(data);
		FileOutputStream fos = new FileOutputStream(path);
		workbook.write(fos);
		fos.close();
		System.out.println("End writing " + data + " into this file path: " + path + "   Sheet name: " + sheetName
				+ "   Column name: " + colName + "");
	}
	
	
	
	

}
