package unitTestPractice;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PracticeTs {

	 public FileInputStream fis = null;
	    public FileOutputStream fos = null;
	    public XSSFWorkbook workbook = null;
	    public XSSFSheet sheet = null;
	    public XSSFRow row = null;
	    public XSSFCell cell = null;
	    String xlFilePath;
	 
	    public PracticeTs(String xlFilePath) throws Exception
	    {
	        this.xlFilePath = xlFilePath;
	        fis = new FileInputStream(xlFilePath);
	        workbook = new XSSFWorkbook(fis);
	        fis.close();
	    }
	    /*public boolean writeResult(String wsName, String colName, int rowNumber, String Result){
			try{
				int sheetIndex=workbook.getSheetIndex(wsName);
				if(sheetIndex==-1)
					return false;			
				int colNum = retrieveNoOfCols(wsName);
				int colNumber=-1;
						
				
				XSSFRow Suiterow = ws.getRow(0);			
				for(int i=0; i<colNum; i++){				
					if(Suiterow.getCell(i).getStringCellValue().equals(colName.trim())){
						colNumber=i;					
					}					
				}
				
				if(colNumber==-1){
					return false;				
				}
				
				XSSFRow Row = ws.getRow(rowNumber);
				XSSFCell cell = Row.getCell(colNumber);
				if (cell == null)
			        cell = Row.createCell(colNumber);			
				
				cell.setCellValue(Result);
				
				opstr = new FileOutputStream(filelocation);
				wb.write(opstr);
				opstr.close();
				
				
			}catch(Exception e){
				e.printStackTrace();
				return false;
			}
			return true;
		}*/
	 
	    public boolean setCellData(String sheetName, int colNumber, int rowNum, String value)
	    {
	        try
	        {
	            sheet = workbook.getSheet(sheetName);
	            row = sheet.getRow(rowNum);
	            if(row==null)
	                row = sheet.createRow(rowNum);
	 
	            cell = row.getCell(colNumber);
	            String colName="Raghu";
	            if(cell == null)
	                cell = row.createCell(colNumber);
	 
	            cell.setCellValue(value);
	 
	            fos = new FileOutputStream(xlFilePath);
	            workbook.write(fos);
	            fos.close();
	        }
	        catch (Exception ex)
	        {
	            ex.printStackTrace();
	            return  false;
	        }
	        return true;
	    }
	    public static void main(String args[]) throws Exception
	    {
	    	PracticeTs eat = new PracticeTs("C:\\Users\\Public\\Documents\\SampleFramework\\TestSuites&TestCases\\New folder\\Test.xlsx");
	        eat.setCellData("Test Data",5,1,"po");
	    }

}
