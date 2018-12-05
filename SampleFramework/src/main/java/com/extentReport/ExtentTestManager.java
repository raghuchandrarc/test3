package com.extentReport;

import java.util.HashMap;
import java.util.Map;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;

        /* OB: extentTestMap holds the information of thread ids and ExtentTest instances.
                ExtentReports instance created by calling getReporter() method from ExtentManager.
                At startTest() method, an instance of ExtentTest created and put into extentTestMap with current thread id.
                At endTest() method, test ends and ExtentTest instance got from extentTestMap via current thread id.
                At getTest() method, return ExtentTest instance in extentTestMap by using current thread id.
         */

public class ExtentTestManager {
	public static ExtentTest test;
    static Map extentTestMap = new HashMap();
    static ExtentReports extent = ExtentManager.getReporter();

    public static synchronized ExtentTest getTest() {
        return (ExtentTest)extentTestMap.get((int) (long) (Thread.currentThread().getId()));
    }

    public static synchronized void endTest() {
        extent.endTest((ExtentTest)extentTestMap.get((int) (long) (Thread.currentThread().getId())));
    }
    

    @SuppressWarnings("unchecked")
	public static synchronized ExtentTest startTest(String testName, String desc) {
    	
        test = extent.startTest(testName, desc);
        
        extent.setTestRunnerOutput(testName);
        
        extentTestMap.put((int) (long) (Thread.currentThread().getId()), test);
      
        return test;
    }
    
}