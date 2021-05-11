package com.tsoft.bot.frontend.helpers.fijas;

import com.tsoft.bot.frontend.listener.Listener;
import com.tsoft.bot.frontend.utility.GenerateWord;
import cucumber.api.Scenario;
import cucumber.api.java.After;
import cucumber.api.java.Before;
import io.appium.java_client.AppiumDriver;
import io.appium.java_client.MobileElement;
import io.appium.java_client.remote.MobileCapabilityType;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.DesiredCapabilities;

import java.io.IOException;

import java.net.URL;
import java.util.concurrent.TimeUnit;

public class HookFijas extends Listener {

    private static final long DELAY = 10;
    public static AppiumDriver<WebElement> driver;
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
        caps.setCapability(MobileCapabilityType.DEVICE_NAME,"SM-J415G");
        caps.setCapability(MobileCapabilityType.UDID,"a84aa87a");
        caps.setCapability(MobileCapabilityType.NEW_COMMAND_TIMEOUT, 120);
        caps.setCapability("appPackage", "com.telefonica.ventafija.dev");
        caps.setCapability("appActivity", "com.telefonica.ventafija.ui.splash.SplashActivity");

        URL url = new URL("http://127.0.0.1:4723/wd/hub");
        driver = new AppiumDriver<WebElement>(url, caps);
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

    public static AppiumDriver<WebElement> getDriver() { return driver; }

}