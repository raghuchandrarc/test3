package com.actions;

/**
 * The ExecuteTestCases class is used to execute the test cases
 *
 * 
 */
import java.awt.event.WindowEvent;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.Assert;
import org.testng.ITest;
import org.testng.ITestResult;
import org.testng.Reporter;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import org.testng.internal.BaseTestMethod;

import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;
import com.driver.MainTestNG;
import com.driver.Driver;
import com.extentReport.ExtentManager;
import com.extentReport.ExtentTestManager;
import com.relevantcodes.extentreports.LogStatus;

import util.ExcelAction;
import util.ExcelLibrary;
import util.ReadConfigProperty;



/**
 * From MainTesnNG, this class will be called
 */
public class ExecuteTestCases implements ITest {

	public static final String FIREFOX = "firefox";
	public static final String IE = "IE";
	public static final String CHROME = "chrome";
	String cbrowserName = "BROWSER_NAME";
	String curl = "URL";
	String cdriver = "IEDRIVER";
	ReadConfigProperty config = new ReadConfigProperty();
	static ExcelLibrary lib = new ExcelLibrary();

	static WebDriver driver = null;
	static ExcelAction act = new ExcelAction();

	static List list;
	protected String mTestCaseName = "";
	/**
	 * In this class, this is the first method to be executed. Reading testsuite
	 * , test case sheet and data sheet and storing the values in Hashmap
	 * @throws IOException 
	 * @throws InvalidFormatException 
	 **/
	@BeforeClass
	public void setup() throws InvalidFormatException, IOException {
		MainTestNG.LOGGER.info("**************************Begin Read Excel Data *********************************");
		System.out.println    (" 																				  ");					
		MainTestNG.LOGGER.info(ExecuteTestCases.class.getName() + "   setup() method called");

		act.readTestSuite();
		act.readTestCaseInExcel();
		act.readTestDataSheet();
		act.readCapturedObjectProperties();
		
		

		/**
		 * Selecting which browser to be executed
		 **/
		String browserName = config.getConfigValues(cbrowserName);
		selectBrowser(browserName);

		MainTestNG.LOGGER.info(config.getConfigValues(curl));
		/**
		 * launching the url
		 **/
		driver.get(config.getConfigValues(curl));
		
		Driver.setDriver(driver);
	}

	/**
	 * Select the browser on which you want to execute tests
	 **/
	private void selectBrowser(String browserName) {

		switch (browserName) {

		case FIREFOX:

			firefoxProfile();
			break;

		case IE:
			ie();
			break;

		case CHROME:
			chrome();
			break;

		default:
			firefoxProfile();
			break;
		}

	}

	private void chrome() {

		System.setProperty("webdriver.chrome.driver", config.getConfigValues(cdriver));

		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	private void ie() {
		DesiredCapabilities returnCapabilities = DesiredCapabilities.internetExplorer();
		returnCapabilities.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS, true);
		returnCapabilities.setCapability(InternetExplorerDriver.ENABLE_PERSISTENT_HOVERING, false);
		System.setProperty("webdriver.ie.driver", config.getConfigValues(cdriver));

		driver = new InternetExplorerDriver(returnCapabilities);
		driver.manage().window().maximize();
		MainTestNG.LOGGER.info("IE launch");

	}

	/**
	 * Firefox profile will help in automatic download of files
	 */
	private void firefoxProfile() {

		FirefoxProfile profile = new FirefoxProfile();
		profile.setPreference("browser.download.folderList", 1);
		profile.setPreference("browser.download.manager.showWhenStarting", false);

		profile.setPreference("browser.helperApps.neverAsk.saveToDisk",
				"application/vnd.openxmlformats-officedocument.wordprocessingml.document");

		driver = new FirefoxDriver(profile);
		driver.manage().window().maximize();

	}

	/**
	 * To override the test case name in the report
	 */
	@BeforeMethod(alwaysRun = true)
	public void testData(Object[] testData) {
		
		String testCase = "";
		if (testData != null && testData.length > 0) {
			String testName = null;
			// Check if test method has actually received required parameters
			for (Object tstname : testData) {
				if (tstname instanceof String) {
					testName = (String) tstname;
					break;
				}
			}
			testCase = testName;
			
			
		}
		this.mTestCaseName = testCase;

	}

	public String getTestName() {
		return this.mTestCaseName;
	}

	public void setTestName(String name) {

		this.mTestCaseName = name;

	}


	/**
	 * All the test cases execution will start from here, which will take the
	 * input from the data provider
	 * @throws Exception 
	 */
	@Test(dataProvider = "dp")
	public void Test(String testName) throws Exception {
		MainTestNG.LOGGER.info("****************************************Begin***********************************************************");
		System.out.println    (" 																				  ");
		System.out.println    ("Reading test case "+testName);
				
		MainTestNG.LOGGER.info(ExecuteTestCases.class.getName() + "  @Test method called" + "    " + testName);

		try {
			this.setTestName(testName);
			
			

			MainTestNG.LOGGER.info("testSuiteIterate() calling");
			
			act.testSuiteIterate(testName);
			Reporter.log("Test case name :" + getTestName());
			MainTestNG.LOGGER.info("****************************************End***********************************************************");
			System.out.println    (" 																				  ");
							
							

		} catch (Exception ex) {
			stack(ex);
			
			Reporter.log("Exception while executing testCase====> " + ex);
			ExtentTestManager.getTest().log(LogStatus.FAIL,ex.getMessage());
			Assert.fail();

		}
		
	}

	@AfterMethod
	public void setResultTestName(ITestResult result) {
		try {
			BaseTestMethod bm = (BaseTestMethod) result.getMethod();
			Field f = bm.getClass().getSuperclass().getDeclaredField("m_methodName");
			f.setAccessible(true);
			//f.set(bm, mTestCaseName);
			
			//f.set(bm,mTestCaseName);
			Reporter.log("Test Case Name :" + bm.getMethodName());
			this.mTestCaseName = " ";

		} catch (Exception ex) {
			stack(ex);
			Reporter.log("exception " + ex.getMessage());
		}
		
	}
	/*@AfterMethod
	public void setResultTestName(ITestResult result) {
		try {
			BaseTestMethod bm = (BaseTestMethod) result.getMethod();
			System.out.println("11111111 "+bm.getMethodName());
			System.out.println("22222222 "+bm.getXmlTest());
			Field f = bm.getClass().getSuperclass().getDeclaredField("m_methodName");
			f.setAccessible(true);
			System.out.println("2222222 "+f.getName());
			//f.set(bm, mTestCaseName);
			
			//f.set(bm.getDate(),mTestCaseName);
			Reporter.log("Test Case Name :" + bm.getMethodName());
			this.mTestCaseName = " ";

		} catch (Exception ex) {
			stack(ex);
			Reporter.log("exception " + ex.getMessage());
		}
		
	}*/
	
	

	/**
	 * Parameterization in testng
	 */
	@DataProvider(name = "dp")
	public Object[][] regression() {
		List list = (ArrayList) ExcelAction.listOfTestCases;
		Object[][] data = new Object[list.size()][1];
		MainTestNG.LOGGER.info(ExecuteTestCases.class.getName() + " TestCases to be executed" + "    " + data);
		for (int i = 0; i < data.length; i++) {
			data[i][0] = (String) list.get(i);
		}
		return data;
	}

	@AfterClass
	public void cleanup() {
		act.clean();
		lib.clean();
	
	}

	public static void stack(Exception e) {
		e.printStackTrace();
		MainTestNG.LOGGER.info(ExecuteTestCases.class.getName() + " Exception occured " + "    " + e.getCause());
	}
}
