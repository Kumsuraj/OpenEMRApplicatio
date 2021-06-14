package com.cyient.page;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.Select;

public class SearchOrAddPatientPage {
	private By patFrameLocator = By.xpath("//iframe[@name='pat']");
	private By fnameLocator = By.xpath("//input[@id='form_fname']");
	private By lnameLocator = By.xpath("//input[@id='form_lname']");
	private By dobLocator = By.xpath("//input[@id='form_DOB']");
	private By selectGenderLocator = By.xpath("//select[@id='form_sex']");
	private By createNewPatientLocator = By.xpath("//button[normalize-space()='Create New Patient']");
	private By modalframeLocator = By.xpath("//iframe[@id='modalframe']");
	private By confirmCreateNewPatientLocator = By.xpath("//input[@value='Confirm Create New Patient']");
	private By closeDlgIframeLocator = By.xpath("//div[@class='closeDlgIframe']");
	private WebDriver driver;

	public SearchOrAddPatientPage(WebDriver driver) {
		this.driver = driver;
	}

	public void switchTOPatFrame() {
		driver.switchTo().frame(driver.findElement(patFrameLocator));
	}

	public void fName(String fname) {
		driver.findElement(fnameLocator).sendKeys(fname);
	}

	public void lName(String lname) {
		driver.findElement(lnameLocator).sendKeys(lname);
	}

	public void dob(String dob) {
		driver.findElement(dobLocator).sendKeys(dob);
	}

	public void selectGender(String dob) {
		Select selectGender = new Select(driver.findElement(selectGenderLocator));
		selectGender.selectByVisibleText(dob);
	}

	public void createNewPatient() {
		driver.findElement(createNewPatientLocator).click();
	}

	public void switchToOutOfFrame() {
		driver.switchTo().defaultContent();
	}

	public void switchTomodalframe() throws InterruptedException {
		Thread.sleep(5000);
		driver.switchTo().frame(driver.findElement(modalframeLocator));
	}

	public void confirmCreateNewPatient() {
		driver.findElement(confirmCreateNewPatientLocator).click();
	}
	
	public String alertsoap() throws InterruptedException {
		String alert = driver.switchTo().alert().getText();
		System.out.println(alert);
		Thread.sleep(3000);
		driver.switchTo().alert().accept();
		return alert;
	}
	
	public void closeDlgIframe() {
		driver.findElement(closeDlgIframeLocator).click();
	}
}
