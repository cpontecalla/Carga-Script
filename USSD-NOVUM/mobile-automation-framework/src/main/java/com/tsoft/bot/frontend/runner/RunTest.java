package com.tsoft.bot.frontend.runner;

import cucumber.api.CucumberOptions;
import cucumber.api.testng.AbstractTestNGCucumberTests;
import org.testng.annotations.Test;


@CucumberOptions(
		features={"src//main//resources//features"},
		glue={"com.tsoft.bot.frontend.steps.APP_MiMovistar","com.tsoft.bot.frontend.helpers",
				"com.tsoft.bot.frontend.steps.APP_USSD", "com.tsoft.bot.frontend.steps.APP_AppVentas"},
		plugin = {"pretty", "html:target/cucumber"},
		tags = {" @USSD_Testing_2,@USSD_Testing_10"},
		monochrome = true
	)

@Test
public class RunTest extends AbstractTestNGCucumberTests{

}
