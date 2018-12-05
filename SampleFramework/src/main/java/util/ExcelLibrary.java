package util;
/**
 * The ExcelLibrary class is used to read the data from the excel sheet
 *
 * 
 */
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Reporter;

import com.driver.MainTestNG;



public class ExcelLibrary {
	 public static XSSFRow row = null;
	 public static XSSFCell cell = null;
	 public static XSSFSheet sheet = null;
	 public static XSSFWorkbook workbook = null;
	static Map<String, Workbook> workbooktable = new HashMap<String, Workbook>();
	public static Map<String, Integer> dict = new Hashtable<String, Integer>();
	public static List list = new ArrayList();
	static ReadConfigProperty config = new ReadConfigProperty();

	/**
	 * To get the excel sheet workbook
	 */
	public static Workbook getWorkbook(String path) {
		Workbook workbook = null;
		if (workbooktable.containsKey(path)) {
			workbook = workbooktable.get(path);
		} else {

			try {
				

				File file = new File(path);

				workbook = WorkbookFactory.create(file);

				workbooktable.put(path, workbook);

			} catch (FileNotFoundException e) {
				e.printStackTrace();
				MainTestNG.LOGGER.info("FileNotFoundException" + e);
			} catch (InvalidFormatException e) {
				e.printStackTrace();
				MainTestNG.LOGGER.info("InvalidFormatException" + e);
			} catch (IOException e) {
				e.printStackTrace();
				MainTestNG.LOGGER.info("IOException" + e);
			} catch (Exception e) {
				e.printStackTrace();
			}

		}
		return workbook;
	}

	/**
	 * To get the number of sheets in excel suite
	 */
	public static List<String> getNumberOfSheetsinSuite(String testPath) {
		List<String> listOfSheets = new ArrayList<String>();

		Workbook workbook = getWorkbook(testPath);

		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			listOfSheets.add(workbook.getSheetName(i));
		}

		return listOfSheets;
	}

	/**
	 * To get the number of sheets in test data sheet
	 */
	public static List<String> getNumberOfSheetsinTestDataSheet(String testPath) {
		List<String> listOfSheets = new ArrayList<String>();

		Workbook workbook = getWorkbook(testPath);
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			if (!(workbook.getSheetName(i)).equalsIgnoreCase(config
					.getConfigValues("TestCase_SheetName"))) {
				listOfSheets.add(workbook.getSheetName(i));

			}
		}
		return listOfSheets;

	}
	

	/**
	 * Get the total rows present in excel sheet
	 */
	public static int getRows(String testSheetName, String pathOfFile)
			throws InvalidFormatException, IOException {
		Workbook workbook = getWorkbook(pathOfFile);
		Reporter.log("getting total number of rows");

		Sheet sheet = workbook.getSheet(testSheetName);

		return sheet.getLastRowNum();

	}

	/**
	 * Get the total columns inside excel sheet
	 */
	public static int getColumns(String testSheetName, String pathOfFile)
			throws InvalidFormatException, IOException {
		Workbook workbook = getWorkbook(pathOfFile);
		Reporter.log("getting total number of columns");
		Sheet sheet = workbook.getSheet(testSheetName);
		return sheet.getRow(0).getLastCellNum();

	}

	/**
	 * Get the column names inside excel sheet
	 */
	@SuppressWarnings("unchecked")
	public static List getColumnNames(String testSheetName, String pathOfFile,
			int j) throws InvalidFormatException, IOException {
		Workbook workbook = getWorkbook(pathOfFile);
		Sheet sheet = workbook.getSheet(testSheetName);

		for (int i = 0; i <= j; i++) {
			if (sheet.getRow(0).getCell(i) != null) {
				list.add(sheet.getRow(0).getCell(i).getStringCellValue()
						.toString());
			}
		}

		return list;

	}

	/**
	 * Get the total number of rows for each column inside excel sheet
	 */
	@SuppressWarnings("unchecked")
	public static void getNumberOfRowsPerColumn(String testSheetName,
			String pathOfFile, int j) throws InvalidFormatException,
			IOException {
		Workbook workbook = getWorkbook(pathOfFile);
		Sheet sheet = workbook.getSheet(testSheetName);
		int totColumns = sheet.getRow(0).getLastCellNum();
		for (int i = 0; i <= totColumns; i++) {
			if (sheet.getRow(0).getCell(i) != null) {
				list.add(sheet.getRow(0).getCell(i).getStringCellValue()
						.toString());
			}
		}
	}

	/**
	 * Read the content of the cell
	 */
	public static String readCell(int rowNum, int colNum, String testSheetName, String pathOfFile) {
		Workbook workbook;
		String cellValue = null;

		workbook = getWorkbook(pathOfFile);
		Sheet sheet = workbook.getSheet(testSheetName);

		Row row = sheet.getRow(rowNum);
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		try {
			if (row != null) {
				Cell cell = row.getCell(colNum);
				if (cell != null) {

					if (cell.getCellTypeEnum() == CellType.STRING)
						return cell.getStringCellValue();
					else if (cell.getCellTypeEnum() == CellType.NUMERIC || cell.getCellTypeEnum() == CellType.FORMULA) {
						cellValue = String.valueOf(cell.getNumericCellValue());
						if (HSSFDateUtil.isCellDateFormatted(cell)) {
							DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
							Date date = cell.getDateCellValue();
							cellValue = df.format(date);
						}
						return cellValue;
					} else if (cell.getCellTypeEnum() == CellType.BLANK)
						return "";
					else
						return String.valueOf(cell.getBooleanCellValue());
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			return "row " + rowNum + " or column " + colNum + " does not exist  in Excel";
		}
		return cellValue;
	}
		
        
								
				//DataFormatter dataFormatter = new DataFormatter();
				//String data = dataFormatter.formatCellValue(cell);
				//cellValue = data;
			
	

	/*public static void writeData(String Path, String sheet, String Testdata, int colNum) throws Exception {
		
		FileInputStream fis = new FileInputStream(Path);

		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheetName = workbook.getSheet(sheet);

		Object[][] datatypes = { { Testdata } };
		//int rowNum = sheetName.getLastRowNum();
		 int rowNum = 0;
		for (Object[] datatype : datatypes) {

			Row row = sheetName.getRow(rowNum++);
			// Row row = sheetName.getRow(rowNum);
			for (Object field : datatype) {
				Cell cell = row.createCell(colNum++);
				if (field instanceof String) {
					cell.setCellValue((String) field);
				} else if (field instanceof Integer) {
					cell.setCellValue((Integer) field);
				} else if (field instanceof Double) {
					cell.setCellValue((Double) field);
				}
			}
		}

		try {
			FileOutputStream outputStream = new FileOutputStream(Path);
			XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
			workbook.write(outputStream);
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println("End writing Data " + Testdata + " into this file path: " + Path + " into this Sheet name: "
				+ sheetName + "");

	}*/
	public static void setcelldata(String path, String sheetName, String colName,  String data)
			throws Exception {
		// create an object of Workbook and pass the FileInputStream object into
		// it to create a pipeline between the sheet and eclipse.
		FileInputStream fis = new FileInputStream(path);
		workbook = new XSSFWorkbook(fis);
		sheet = workbook.getSheet(sheetName);
		int col_Num = -1;
		row = sheet.getRow(0);
		for (int i = 0; i < row.getLastCellNum(); i++) {
			if (row.getCell(i).getStringCellValue().trim().equals(colName)) {
				col_Num = i;
			}
		}
		//int rowNum =ExcelLibrary.getRowCount(sheetName);
		int rowNum =1;

		//sheet.autoSizeColumn(col_Num);
		row = sheet.getRow(rowNum);
		if (row == null)
			row = sheet.createRow(rowNum);

		cell = row.getCell(col_Num);
		if (cell == null)
			cell = row.createCell(col_Num);
		cell.setCellType(cell.CELL_TYPE_STRING);
		cell.setCellValue(data);
		FileOutputStream fos = new FileOutputStream(path);
		workbook.write(fos);
		fos.close();
		System.out.println("End writing " + data + " into this file path: "+ path +"   Sheet name: " + sheetName +"   Column name: " + colName +"");
	}

	public static int getRowCount(String SheetName) {
		int iNumber = 0;
		try {
			sheet = workbook.getSheet(SheetName);
			iNumber = sheet.getLastRowNum() + 1;
		} catch (Exception e) {
		}
		return iNumber;
	}
	public static void writeXLSFile(String Path, String sheetName, String data) throws IOException {

		// String excelFileName =
		// "C:\\Users\\Public\\Documents\\practiceWorkspace\\Practice\\Excelpath\\Test.xlsx";//name
		// of excel file
		FileInputStream fis = new FileInputStream(Path);

		// String sheetName = "rc";//name of sheet

		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet(sheetName);
		//int rowNum = ExcelLibrary.getRowCount(sheetName);

		// iterating r number of rows
		for (int r = 1; r < 2; r++) {
			XSSFRow row = sheet.createRow(r);

			// iterating c number of columns
			for (int c = 0; c < 1; c++) {
				XSSFCell cell = row.createCell(c);

				cell.setCellValue(data);
			}
		}

		FileOutputStream fileOut = new FileOutputStream(Path);

		// write this workbook to an Outputstream.
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
	}

	/**
	 * To clear the worktable and list
	 */
	public void clean() {
		workbooktable.clear();
		list.clear();
	}

}