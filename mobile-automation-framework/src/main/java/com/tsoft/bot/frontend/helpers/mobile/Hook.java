package com.tsoft.bot.frontend.helpers.mobile;

import com.tsoft.bot.frontend.listener.Listener;
import com.tsoft.bot.frontend.utility.GenerateWord;
import cucumber.api.Scenario;
import cucumber.api.java.After;
import cucumber.api.java.Before;
import io.appium.java_client.AppiumDriver;
import io.appium.java_client.MobileElement;
import io.appium.java_client.remote.MobileCapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;

import java.io.IOException;

import java.net.URL;
import java.util.concurrent.TimeUnit;

public class Hook extends Listener {

	private static final long DELAY = 10;
	public static AppiumDriver<MobileElement> driver;
	private static GenerateWord generateWord = new GenerateWord();

	@Before
	public void Scenario(Scenario scenario){
		onTestStart(scenario.getName());
	}

	@Before
	public void setUpAppium() throws Throwable {
        DesiredCapabilities caps = new DesiredCapabilities();
        caps.setCapability(MobileCapabilityType.PLATFORM_NAME,"Android");
        caps.setCapability(MobileCapabilityType.PLATFORM_VERSION,"9");
        caps.setCapability(MobileCapabilityType.DEVICE_NAME,"MRD-LX3");
        caps.setCapability(MobileCapabilityType.UDID,"5DNNW19412010776");
        caps.setCapability(MobileCapabilityType.NEW_COMMAND_TIMEOUT, 120);
        caps.setCapability("appPackage", "pe.vasslatam.movistar.mobile.sales");
        caps.setCapability("appActivity", "pe.vasslatam.movistar.mobile.sales.activities.SplashActivity");

        URL url = new URL("http://127.0.0.1:4723/wd/hub");
        driver = new AppiumDriver<MobileElement>(url, caps);
        driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
        generateWord.startUpWord();
	}

	@After
	public void tearDown() throws IOException {

		onFinish();
		generateWord.endToWord();

		//getDriver().quit();
        //driver.quit();
	}

	public static AppiumDriver<MobileElement> getDriver() { return driver; }

}