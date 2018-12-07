package com.driver;

/**
 * The MainTestNG class is used to execute the jar
 *
 * 
 */
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.FileHandler;
import java.util.logging.Formatter;
import java.util.logging.Handler;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

import org.testng.TestNG;
import org.testng.xml.XmlClass;
import org.testng.xml.XmlSuite;
import org.testng.xml.XmlTest;

import util.ReadConfigProperty;

/**
 * This is the class from where the execution gets started.
 */

// TestSuite
public class MainTestNG {
	static Handler filehandler;
	static Formatter formatter = null;
	public static final Logger LOGGER = Logger.getLogger(MainTestNG.class.getName());
	ReadConfigProperty config = new ReadConfigProperty();
	

	static String dir = "user.dir";

	public static void main(String[] args) {
		try {

			filehandler = new FileHandler("./log.txt");
		} catch (SecurityException e) {
			MainTestNG.LOGGER.info(e.getMessage());
		} catch (IOException e) {
			MainTestNG.LOGGER.info(e.getMessage());
		}

		LOGGER.addHandler(filehandler);
		formatter = new SimpleFormatter();
		filehandler.setFormatter(formatter);
		LOGGER.info("Logger Name: " + LOGGER.getName());
		ReadConfigProperty.configpath = System.getProperty(dir);
		MainTestNG test = new MainTestNG();

		/**
		 * testNG execution starts here
		 */
		test.testng();
	}

	/**
	 * adding listners, setting test-output folder Mentioning the TestSuite Name
	 */
	public void testng() {
		// RegressionSuite
		
		TestNG objTestNG = new TestNG();
		XmlSuite TSuite = new XmlSuite();
		TSuite.setName("Test Suite");
		TSuite.addListener("org.uncommons.reportng.HTMLReporter");
		TSuite.addListener("org.uncommons.reportng.JUnitXMLReporter");
		TSuite.addListener("util.TestListener");
		objTestNG.setOutputDirectory("test-output");
		XmlTest myTest = new XmlTest(TSuite);
		myTest.setName(" Test Suite Begin...");
		List<XmlClass> myClasses = new ArrayList<XmlClass>();
		myClasses.add(new XmlClass("com.actions.ExecuteTestCases"));
		myTest.setXmlClasses(myClasses);
		List<XmlTest> myTests = new ArrayList<XmlTest>();
		myTests.add(myTest);
		TSuite.setTests(myTests);
		List<XmlSuite> mySuites = new ArrayList<XmlSuite>();
		mySuites.add(TSuite);
		objTestNG.setXmlSuites(mySuites);
		objTestNG.setUseDefaultListeners(true);
		objTestNG.run();

	}

}
