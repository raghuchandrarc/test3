package com.driver;
/**
 * The WebDriverClass class is used to generate a single instance of driver to
 * launch the browser
 *
 * 
 */
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;

public class Driver {
	static WebDriver driver;

	private Driver() {
		Driver.getDriver();
	}

	public static WebDriver getDriver() {
		driver.manage().timeouts().pageLoadTimeout(20, TimeUnit.SECONDS);
		
		return driver;
	}

	public static void setDriver(WebDriver driver) {
		Driver.driver = driver;
	}

	/**
	 * @return
	 * Getting the instance of the driver
	 */
	public static WebDriver getInstance() {
		if (driver == null) {

			driver = (WebDriver) new Driver();

			return driver;
		} else {

			return driver;
		}

	}

}
