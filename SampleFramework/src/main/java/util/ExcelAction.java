package util;

/**
 * The ExcelAction class is used to store the data from the excel into map
 *
 * 
 */
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.WebDriver;
import org.testng.Reporter;

import com.actions.MethodType;
import com.driver.MainTestNG;
import com.driver.Driver;
import com.extentReport.ExtentTestManager;
import com.model.CapturedObjectPropModel;
import com.model.TestCase;
import com.relevantcodes.extentreports.LogStatus;

public class ExcelAction {
	WebDriver driver;
	static ExcelLibrary excel = new ExcelLibrary();
	static ReadConfigProperty config = new ReadConfigProperty();
	static Map<String, Object> testCaseSheet = new HashMap<String, Object>();

	static Map<String, String> readFromConfigFile = new HashMap<String, String>();
	static Map<String, Object> testSuiteSheet = new HashMap<String, Object>();
	static Map<String, Object> testDataSheet = new HashMap<String, Object>();
	static Map<String, Object> capObjPropSheet = new HashMap<String, Object>();

	public static List listOfTestCases = new ArrayList();
	int numberOfTimeExecution = 0;
	MethodType methodtype = new MethodType();

	String testcasepth = "TestCasePath";
	static String actionType = "";
	static String objectLocators = "";

	/*
	 * public static void main(String[] args) { ExcelAction action = new
	 * ExcelAction(); action.readCapturedObjectProperties();
	 * action.readLocators("PAGE", "SEARCH_BOX"); }
	 */

	/**
	 * Read test data sheet
	 */
	@SuppressWarnings("unchecked")
	public void readTestDataSheet() {

		String sheetName;
		String pathOFFile = config.getConfigValues(testcasepth);
		List<String> list = ExcelLibrary.getNumberOfSheetsinTestDataSheet(config.getConfigValues(testcasepth));
		for (int i = 0; i < list.size(); i++) {
			sheetName = list.get(i);
			Map<String, Object> temp1 = new HashMap<String, Object>();

			try {
				Reporter.log("sheetName" + sheetName + "----" + "sheetName, pathOFFile" + pathOFFile);
				List listColumnNames = ExcelLibrary.getColumnNames(sheetName, pathOFFile,
						ExcelLibrary.getColumns(sheetName, pathOFFile));
				// iterate through columns in sheet
				for (int j = 0; j < listColumnNames.size(); j++) {
					// get Last Row for each Column
					int row = 1;
					List listColumnValues = new ArrayList();
					do {
						listColumnValues.add(ExcelLibrary.readCell(row, j, sheetName, pathOFFile));
						row++;
					} while ((ExcelLibrary.readCell(row, j, sheetName, pathOFFile)) != null);
					temp1.put((String) listColumnNames.get(j), listColumnValues);
				}
				listColumnNames.clear();
			} catch (InvalidFormatException | IOException e) {
				// check after run
				MainTestNG.LOGGER.info("InvalidFormatException,IOException" + e);
			}
			testDataSheet.put(sheetName, temp1);
		}
	}

	/**
	 * Iterate over each row in testcase sheet and pass the data to execute
	 * method in MethodType.java
	 */
	public void testSuiteIterate(String tcName) {
		MainTestNG.LOGGER.info("testSuiteIterate() called  " + tcName);

		String key = tcName;
		
		TestCase temp = (TestCase) testCaseSheet.get(key);
		
		
		List testStepId = temp.getTestStepId();
		
		System.out.println("Begin of Testcase ID == "+testStepId);
		
		
		Reporter.log("<=========================Begin Of TestStep=========================>");

		List dataColValues = null;
		int noOfExecution = 0;
		for (int i = 0; i < testStepId.size(); i++) {
			if (!(temp.getTestData().get(i).isEmpty())) {
				if (temp.getTestData().get(i).contains(".")) {

					String data = temp.getTestData().get(i);
					String[] testDataArray = data.split("\\.");

					dataColValues = getColumnValue(testDataArray);

					noOfExecution = dataColValues.size();

					break;
				}
			} else {
				noOfExecution = 0;
			}
		}
		MainTestNG.LOGGER.info("columnValue added newly numberOfTimesExecution===" + dataColValues);
		MainTestNG.LOGGER.info("testCaseExecution==" + noOfExecution);

		ExtentTestManager.startTest(tcName, temp.getTestCaseName());

		if (noOfExecution != 0) {
			for (int execution = 0; execution < noOfExecution; execution++) {
				for (int i = 0; i < testStepId.size(); i++) {

					String methodType = temp.getMethodType().get(i);
					Reporter.log("Page ==============================>" + methodType);
					
					objectLocators = temp.getObjectNameFromPropertiesFile().get(i);
					Reporter.log("Object or Locators ================>" + objectLocators);
					// String actionType = temp.getActionType().get(i);
					actionType = temp.getActionType().get(i);

					Reporter.log("Action Keyword =====================>" + actionType);
					

					// Data Sheet logic
					if (!(temp.getTestData().get(i).isEmpty())) {

						if (temp.getTestData().get(i).contains(".")) {

							String data = temp.getTestData().get(i);
							String[] testDataArray = data.split("\\.");

							List InputData = getColumnValue(testDataArray);

							Reporter.log("InputData =====================>" + InputData);

							try {
								Reporter.log("testCaseExecution ====>" + noOfExecution);

								Reporter.log("																	");
								Reporter.log("																	");
								List<String> list = readLocators(methodType, objectLocators);
								
								
								String ObjectName = objectLocators;
								methodType = list.get(0);
								objectLocators = list.get(1);
								MainTestNG.LOGGER.info("methodType=" + methodType);
								MainTestNG.LOGGER.info("objectLocators as name=" + objectLocators);

								methodtype.methodExecutor(methodType, objectLocators, actionType,
										InputData.get(execution).toString());

								ExtentTestManager.getTest().log(LogStatus.PASS, ObjectName + "==========> " + actionType
										+ "==========>  " + objectLocators + "==========>  " + InputData.toString());

							}

							catch (IndexOutOfBoundsException e) {

								String s = e.getMessage();
								throw new IndexOutOfBoundsException(
										"data column is blank..Please enter value in datasheet" + s);

							}

						}

						if (execution == noOfExecution) {

							break;
						}
					} else {
						driver = Driver.getInstance();

						List<String> list = readLocators(methodType, objectLocators);
						
						methodType = list.get(0);
						
						objectLocators = list.get(1);
						MainTestNG.LOGGER.info("methodType=" + methodType);
						methodtype.methodExecutor(methodType, objectLocators, actionType, null);
					}
				}
				if (execution == noOfExecution) {
					break;
				}
			}

		} else {
			for (int i = 0; i < testStepId.size(); i++) {

				String methodType = temp.getMethodType().get(i);
				String objectLocators = temp.getObjectNameFromPropertiesFile().get(i);
				String actionType = temp.getActionType().get(i);

				driver = Driver.getInstance();

				List<String> list = readLocators(methodType, objectLocators);
				methodType = list.get(0);
				objectLocators = list.get(1);
				MainTestNG.LOGGER.info("methodType=" + methodType);
				MainTestNG.LOGGER.info("objectLocators=" + objectLocators);

				methodtype.methodExecutor(methodType, objectLocators, actionType, null);
			}

		}

	}

	private List getColumnValue(String[] testDataArray) {

		@SuppressWarnings("unchecked")
		Map<String, Object> dataSheet = (HashMap<String, Object>) testDataSheet.get(testDataArray[0]);
		List coulmnValue = (ArrayList) dataSheet.get(testDataArray[1]);
		return coulmnValue;
		
	}

	/**
	 * populate data to testSuitedata and listOfTestCases to be executed
	 */
	@SuppressWarnings("unchecked")
	public void readTestSuite() {
		readFromConfigFile = config.readConfigFile();

		for (String suiteName : readFromConfigFile.values()) {

			String testSuiteFilePath = config.getConfigValues("TestSuiteName");
			System.out.println(testSuiteFilePath);
			List<String> suiteSheets = ExcelLibrary.getNumberOfSheetsinSuite(testSuiteFilePath);
			System.out.println(suiteSheets.size());

			for (int i = 0; i < suiteSheets.size(); i++) {
				String sheetName = suiteSheets.get(i);
				System.out.println(sheetName);
				if (suiteName.trim().equalsIgnoreCase(sheetName)) {
					Map<String, Object> temp1 = new HashMap<String, Object>();
					try {
						for (int row = 1; row <= ExcelLibrary.getRows(sheetName, testSuiteFilePath); row++) {

							String testCaseName = ExcelLibrary.readCell(row, 0, suiteName.trim(), testSuiteFilePath);
							String testCaseState = ExcelLibrary.readCell(row, 1, suiteName.trim(), testSuiteFilePath);

							if (("YES").equalsIgnoreCase(testCaseState)) {
								listOfTestCases.add(testCaseName);
							}
							temp1.put(testCaseName, testCaseState);

						}
						Reporter.log("listOfTestCases=============*****************" + listOfTestCases);
						testSuiteSheet.put(suiteName, temp1);
					} catch (InvalidFormatException | IOException e) {

						MainTestNG.LOGGER.info("e" + e);

					}
				}
			}
		}

	}

	/**
	 * Read the content of the excel testcase sheet and store the data in model
	 * and store this model in hashmap
	 * @throws InvalidFormatException 
	 * @throws IOException 
	 */
	public void readTestCaseInExcel() throws InvalidFormatException, IOException {

		String testsheetnme = "TestCase_SheetName";
		String testCasePath = config.getConfigValues(testcasepth);
		String testCaseSheetName = config.getConfigValues(testsheetnme);
		String Reuse="Reuse_";

		TestCase tc = null;
		try {
			for (int row = 1; row <= ExcelLibrary.getRows(testCaseSheetName, testCasePath); row++) {

				if (!(ExcelLibrary.readCell(row, 0, testCaseSheetName, testCasePath).isEmpty())) {

					tc = new TestCase();
					
					tc.setTestCaseName(ExcelLibrary.readCell(row, 0, testCaseSheetName, testCasePath));
					
					String ReusetestCaseSheetName =tc.getTestCaseName();
					String TestCaseSheetName =tc.getTestCaseName();
					System.out.println("Running Test cases " +TestCaseSheetName);
				if(ReusetestCaseSheetName.startsWith(Reuse)){
					for (int row1 = 1; row1 <= ExcelLibrary.getRows(ReusetestCaseSheetName, testCasePath); row1++) {
						if (!(ExcelLibrary.readCell(row1, 0, ReusetestCaseSheetName, testCasePath).isEmpty())) {
							tc.setTestCaseName(ExcelLibrary.readCell(row1, 0, ReusetestCaseSheetName, testCasePath));

							tc.setTestStepId(ExcelLibrary.readCell(row1, 1, ReusetestCaseSheetName, testCasePath));
						
							

							tc.setMethodType(ExcelLibrary.readCell(row1, 3, ReusetestCaseSheetName, testCasePath));
							
							
							tc.setObjectNameFromPropertiesFile(ExcelLibrary.readCell(row1, 4, ReusetestCaseSheetName, testCasePath));
							tc.setActionType(ExcelLibrary.readCell(row1, 5, ReusetestCaseSheetName, testCasePath));

							tc.setTestData(ExcelLibrary.readCell(row1, 6, ReusetestCaseSheetName, testCasePath));
							testCaseSheet.put(ExcelLibrary.readCell(row1, 0, ReusetestCaseSheetName, testCasePath), tc);
							
						}
						else{
							tc.setTestStepId(ExcelLibrary.readCell(row1, 1, ReusetestCaseSheetName, testCasePath));
							tc.setMethodType(ExcelLibrary.readCell(row1, 3, ReusetestCaseSheetName, testCasePath));
							
							
							tc.setObjectNameFromPropertiesFile(ExcelLibrary.readCell(row1, 4, ReusetestCaseSheetName, testCasePath));
							tc.setActionType(ExcelLibrary.readCell(row1, 5, ReusetestCaseSheetName, testCasePath));

							tc.setTestData(ExcelLibrary.readCell(row1, 6, ReusetestCaseSheetName, testCasePath));
							
						}
						
					}
					
						//readReuseTestCaseInExcel(testsheetnme, testCasePath, ReusetestCaseSheetName);
					
				}
				else {
					try {
					if(TestCaseSheetName.equalsIgnoreCase(TestCaseSheetName)){
						for (int row1 = 1; row1 <= ExcelLibrary.getRows(TestCaseSheetName, testCasePath); row1++) {
							if (!(ExcelLibrary.readCell(row1, 0, TestCaseSheetName, testCasePath).isEmpty())) {
								tc.setTestCaseName(ExcelLibrary.readCell(row1, 0, TestCaseSheetName, testCasePath));

								tc.setTestStepId(ExcelLibrary.readCell(row1, 1, TestCaseSheetName, testCasePath));
							
								

								tc.setMethodType(ExcelLibrary.readCell(row1, 3, TestCaseSheetName, testCasePath));
								
								
								tc.setObjectNameFromPropertiesFile(ExcelLibrary.readCell(row1, 4, TestCaseSheetName, testCasePath));
								tc.setActionType(ExcelLibrary.readCell(row1, 5, TestCaseSheetName, testCasePath));

								tc.setTestData(ExcelLibrary.readCell(row1, 6, TestCaseSheetName, testCasePath));
								testCaseSheet.put(ExcelLibrary.readCell(row1, 0, TestCaseSheetName, testCasePath), tc);
								
							}
							else{
								tc.setTestStepId(ExcelLibrary.readCell(row1, 1, TestCaseSheetName, testCasePath));
								tc.setMethodType(ExcelLibrary.readCell(row1, 3, TestCaseSheetName, testCasePath));
								
								
								tc.setObjectNameFromPropertiesFile(ExcelLibrary.readCell(row1, 4, TestCaseSheetName, testCasePath));
								tc.setActionType(ExcelLibrary.readCell(row1, 5, TestCaseSheetName, testCasePath));

								tc.setTestData(ExcelLibrary.readCell(row1, 6, TestCaseSheetName, testCasePath));
								
							}
							
						}
						
							//readReuseTestCaseInExcel(testsheetnme, testCasePath, ReusetestCaseSheetName);
						
					}
					
				}
					catch (InvalidFormatException e) {

						MainTestNG.LOGGER.info(e.getMessage());
						throw new InvalidFormatException(
								"Invalid format in test case sheet " + e);
					} catch (IOException e) {

						MainTestNG.LOGGER.info(e.getMessage());
					}
				}
				
					/*tc.setTestStepId(ExcelLibrary.readCell(row, 1, testCaseSheetName, testCasePath));
				
					

					tc.setMethodType(ExcelLibrary.readCell(row, 3, testCaseSheetName, testCasePath));
					
					
					tc.setObjectNameFromPropertiesFile(ExcelLibrary.readCell(row, 4, testCaseSheetName, testCasePath));
					tc.setActionType(ExcelLibrary.readCell(row, 5, testCaseSheetName, testCasePath));

					tc.setTestData(ExcelLibrary.readCell(row, 6, testCaseSheetName, testCasePath));
					testCaseSheet.put(ExcelLibrary.readCell(row, 0, testCaseSheetName, testCasePath), tc);*/
					
				

				
				}/*else {

					tc.setTestStepId(ExcelLibrary.readCell(row, 1, testCaseSheetName, testCasePath));
					tc.setMethodType(ExcelLibrary.readCell(row, 3, testCaseSheetName, testCasePath));
					
					
					tc.setObjectNameFromPropertiesFile(ExcelLibrary.readCell(row, 4, testCaseSheetName, testCasePath));
					tc.setActionType(ExcelLibrary.readCell(row, 5, testCaseSheetName, testCasePath));

					tc.setTestData(ExcelLibrary.readCell(row, 6, testCaseSheetName, testCasePath));
				}*/
			}
		} catch (InvalidFormatException e) {

			MainTestNG.LOGGER.info(e.getMessage());
			throw new InvalidFormatException(
					"Invalid format in test case sheet " + e);
		} catch (IOException e) {

			MainTestNG.LOGGER.info(e.getMessage());
			throw new IOException(
					"Io exception " + e);
		}
	}

	public void clean() {
		excel.clean();

	}

	

	/**
	 * Capture object properties in excel sheet
	 */
	public void readCapturedObjectProperties() {
		String testSheetName = "CapturedObjectProperties";
		String testCasePath = config.getConfigValues(testcasepth);
		MainTestNG.LOGGER.info("testCasePath==" + testCasePath);
		MainTestNG.LOGGER.info("**************************End Read Excel Data *********************************");
		System.out.println(" 																				  ");

		try {
			int totrows = ExcelLibrary.getRows(testSheetName, testCasePath);
			MainTestNG.LOGGER.info("total rows= " + totrows);
			String prevPagename = "";
			Map<String, Object> pageInfo = null;
			for (int j = 1; j <= totrows; j++) {
				String pagename = ExcelLibrary.readCell(j, 0, testSheetName, testCasePath);

				if (prevPagename.equals(pagename)) {

					String page = ExcelLibrary.readCell(j, 0, testSheetName, testCasePath);
					String name = ExcelLibrary.readCell(j, 1, testSheetName, testCasePath);
					String property = ExcelLibrary.readCell(j, 2, testSheetName, testCasePath);
					String value = ExcelLibrary.readCell(j, 3, testSheetName, testCasePath);

					CapturedObjectPropModel capModel = new CapturedObjectPropModel();
					capModel.setPage(page);
					capModel.setName(name);
					capModel.setProperty(property);
					capModel.setValue(value);
					MainTestNG.LOGGER.info(capModel.getPage() + "  " + capModel.getName() + "  " + capModel.getValue()
							+ "  " + capModel.getProperty());
					pageInfo.put(name, capModel);

				} else {
					if (prevPagename != null) {
						capObjPropSheet.put(prevPagename, pageInfo);
					}
					pageInfo = new HashMap<String, Object>();
					String page = ExcelLibrary.readCell(j, 0, testSheetName, testCasePath);
					String name = ExcelLibrary.readCell(j, 1, testSheetName, testCasePath);
					String property = ExcelLibrary.readCell(j, 2, testSheetName, testCasePath);
					String value = ExcelLibrary.readCell(j, 3, testSheetName, testCasePath);

					CapturedObjectPropModel capModel = new CapturedObjectPropModel();
					capModel.setPage(pagename);
					capModel.setName(name);
					capModel.setProperty(property);
					capModel.setValue(value);
					pageInfo.put(name, capModel);
					prevPagename = pagename;
				}

				if (prevPagename != null) {
					capObjPropSheet.put(prevPagename, pageInfo);
				}
			}

		} catch (InvalidFormatException e) {

			MainTestNG.LOGGER.info("InvalidFormatException=" + e);

		} catch (IOException e) {

			e.printStackTrace();
		}
	}

	/**
	 * Capture object Locators in excel sheet
	 */
	public List<String> readLocators(String page, String name) {
		
		MainTestNG.LOGGER.info(page);
		MainTestNG.LOGGER.info(name);
		Map<String, Object> temp = (Map<String, Object>) capObjPropSheet.get(page);
		List<String> locators = new ArrayList<>();

		MainTestNG.LOGGER.info("objects" + capObjPropSheet.get(page));
		if (capObjPropSheet.get(page) != null) {

			MainTestNG.LOGGER.info("name" + temp.get(name));
			CapturedObjectPropModel c = (CapturedObjectPropModel) temp.get(name);
			MainTestNG.LOGGER.info(c.getName());
			MainTestNG.LOGGER.info("c.getPage()=" + c.getPage());

			if (c.getPage().equals(page) && c.getName().equals(name)) {
				locators.add(c.getProperty());
				locators.add(c.getValue());
				
				MainTestNG.LOGGER.info("locators" + locators);
			}
		}
		MainTestNG.LOGGER.info("size " + locators.size());
		return locators;
	}
}
