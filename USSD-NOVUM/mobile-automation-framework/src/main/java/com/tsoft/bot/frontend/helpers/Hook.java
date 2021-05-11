package com.tsoft.bot.frontend.helpers;

import com.tsoft.bot.frontend.listener.Listener;
import com.tsoft.bot.frontend.utility.FileHelper;
import com.tsoft.bot.frontend.utility.GenerateWord;
import cucumber.api.Scenario;
import cucumber.api.java.After;
import cucumber.api.java.Before;
import io.appium.java_client.AppiumDriver;
import io.appium.java_client.MobileElement;
import io.appium.java_client.remote.MobileCapabilityType;
import org.openqa.selenium.Capabilities;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.internal.TouchAction;
import org.openqa.selenium.remote.DesiredCapabilities;


import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.concurrent.TimeUnit;


public class Hook extends Listener {

//	private static final String URL_MOVISTAR_FIJA   = "http://tdp-web-venta-fija-qa.mybluemix.net/acciones";
//    private static final String CHROME_DRIVER = "/src/main/resources/driver/chromedriver.exe";

	public static AppiumDriver<MobileElement> driver;
	static GenerateWord generateWord = new GenerateWord();

	@Before
	public void Scenario(Scenario scenario){
		onTestStart(scenario.getName());
	}

	@Before
	public void setUpAppium() throws Throwable {
		DesiredCapabilities caps = new DesiredCapabilities();
		//caps.setCapability("platformName", "ANDROID");
		//caps.setCapability(CapabilityType.PLATFORM_NAME, "ANDROID");
		caps.setCapability(MobileCapabilityType.PLATFORM_NAME, "Android");
		caps.setCapability(MobileCapabilityType.PLATFORM_VERSION, "5.1.1");
		caps.setCapability(MobileCapabilityType.DEVICE_NAME,"SM J320M");
		caps.setCapability(MobileCapabilityType.UDID, "42009d83a8a61400");
		caps.setCapability(MobileCapabilityType.NEW_COMMAND_TIMEOUT, 120);
		//caps.setCapability(MobileCapabilityType.APP,"C://Users//admin//Desktop//appium-demo//src//test//resources//apps//Mi Movistar.apk");
		//caps.setCapability(MobileCapabilityType.BROWSER_NAME, "Chrome");
		//MI MOVISTAR
		caps.setCapability("appPackage", "tdp.app.col.enterprise"); //Mi Movistar
		caps.setCapability("appActivity", "com.tuenti.messenger.ui.activity.MainActivity"); //Mi Movistar
		//APP VENTAS TESTING
//		caps.setCapability("appPackage", "pe.vasslatam.movistar.mobile.sales.debug"); //APP VENTAS
//		caps.setCapability("appActivity", "pe.vasslatam.movistar.mobile.sales.activities.SplashActivity"); // APP VENTAS
		//USSD
//		caps.setCapability("appPackage", "com.android.contacts"); //USSD
//		caps.setCapability("appActivity", "com.android.contacts.activities.PeopleActivity"); // USSD


		URL url = new URL("http://127.0.0.1:4723/wd/hub");

		driver = new AppiumDriver<MobileElement>(url, caps);
		//driver = new AndroidDriver<MobileElement>(url,caps);
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		generateWord.startUpWord();
	}

	@After
	public void tearDown() throws IOException {
		driver.quit();
		onFinish();
		generateWord.endToWord();
	}

	public static AppiumDriver<MobileElement> getDriver()
	{
		return driver;
	}

}
