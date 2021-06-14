package com.cyient.test;

import org.testng.Assert;
import org.testng.annotations.Test;

import com.cyient.base.WebDriverWrapper;
import com.cyient.page.DashboardPage;
import com.cyient.page.LoginPage;
import com.cyient.page.PatientDashboardPage;
import com.cyient.page.PatientFinderPage;
import com.cyient.page.SearchOrAddPatientPage;

public class PatientTest extends WebDriverWrapper {

	@Test
	public void addPatientTest() throws InterruptedException {

		LoginPage login = new LoginPage(driver);
		login.sendUsername("admin");
		login.sendPassword("pass");
		login.setLanguageByText("English (Indian)");
		login.clickOnLogin();

		// dashboardPage
		DashboardPage dashboard = new DashboardPage(driver);
		dashboard.mouseHoverOnPatientClient();
		dashboard.clickOnPatient();

		// patient finder(class)
		PatientFinderPage addPatient = new PatientFinderPage(driver);
		addPatient.switchToFinFrame();
		addPatient.addNewPatient();
		addPatient.switchToOutOfFinFrame();

		// searchOrAddPatientPage
		SearchOrAddPatientPage soappage = new SearchOrAddPatientPage(driver);
		soappage.switchTOPatFrame();
		soappage.fName("Suraj");
		soappage.lName("Kumar");
		soappage.lName("2021-04-03");
		soappage.selectGender("Male");
		soappage.createNewPatient();
		soappage.switchToOutOfFrame();
		Thread.sleep(5000);
		soappage.switchTomodalframe();
		soappage.confirmCreateNewPatient();
		soappage.switchToOutOfFrame();
		Thread.sleep(5000);
		soappage.alertsoap();
		soappage.switchToOutOfFrame();

		// patientDashboard
		PatientDashboardPage pdPage = new PatientDashboardPage(driver);
		pdPage.switchTOPatFrame();
		pdPage.medicalRecordDash();

		Assert.assertTrue(soappage.alertsoap().contains("New Due Clinical Reminders"));
		Assert.assertEquals(pdPage.medicalRecordDash(), "Medical Record Dashboard - Suraj Kumar");

	}
}
