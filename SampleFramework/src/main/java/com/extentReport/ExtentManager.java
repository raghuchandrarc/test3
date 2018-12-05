package com.extentReport;

import java.io.File;

import com.relevantcodes.extentreports.ExtentReports;

import util.ReadConfigProperty;


//OB: ExtentReports extent instance created here. That instance can be reachable by getReporter() method.

public class ExtentManager {

    public static ExtentReports extent;
    static ReadConfigProperty config = new ReadConfigProperty();
    static String curl = "URL";
    public synchronized static ExtentReports getReporter(){
        if(extent == null){
            //Set HTML reporting file location
            String workingDir = System.getProperty("user.dir");
            extent = new ExtentReports(workingDir+"\\SCREENSHOT\\ExtentReportResults.html", true);
            extent.addSystemInfo("Application URL", config.getConfigValues(curl));
            extent.addSystemInfo("Automation Test Tool", "Selenium 3.5.0");
           
            extent.loadConfig(new File(System.getProperty("user.dir")+"/test-output/extent-config.xml"));
            
        }
        return extent;
    }

	
}
