package com.cyient.test;

import org.openqa.selenium.By;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.cyient.base.WebDriverWrapper;
import com.cyient.page.DashboardPage;
import com.cyient.page.LoginPage;
import com.cyient.utilities.DataProviderUtils;

public class LoginTest extends WebDriverWrapper {

	@Test(dataProvider = "inValidCredentialData", dataProviderClass = DataProviderUtils.class )
	public void invalidCredentialTest(String username, String password, String languageText, String expectedValue) {

		LoginPage login = new LoginPage(driver);
		login.sendUsername(username);
		login.sendPassword(password);
		login.setLanguageByText(languageText);
		login.clickOnLogin();

		Assert.assertEquals(login.getErrorMessage(), 	expectedValue);

	}

	// create dataProvider and connect with validCredential test method

	@Test(dataProvider = "validCredentialExcelData", dataProviderClass = DataProviderUtils.class)
	public void validCredentialTest(String username, String password, String languageText, String expectedValue) {

		LoginPage login = new LoginPage(driver);
		login.sendUsername(username);
		login.sendPassword(password);
		login.setLanguageByText(languageText);
		login.clickOnLogin();

		DashboardPage dashboard = new DashboardPage(driver);
		String actualValue = dashboard.getDashboardPageTitle();

		Assert.assertEquals(actualValue, expectedValue);

	}

	@Test
	public void validateUIComponentTest() {

		Assert.assertEquals(driver.findElement(By.xpath("//p[contains(text(),'most ')]")).getText(),
				"The most popular open-source Electronic Health Record and Medical Practice Management solution.");
		Assert.assertEquals(driver.findElement(By.partialLinkText("Acknowledgments")).getText(),
				"Acknowledgments, Licensing and Certification");
		Assert.assertEquals(driver.findElement(By.id("authUser")).getAttribute("placeholder"), "Username:");
		Assert.assertEquals(driver.findElement(By.id("clearPass")).getAttribute("placeholder"), "Password:");
		Assert.assertEquals(
				driver.findElement(By.xpath("//option[contains(text(),'Default - English (Standard)')]")).getText(),
				"Default - English (Standard)");
	}
}
