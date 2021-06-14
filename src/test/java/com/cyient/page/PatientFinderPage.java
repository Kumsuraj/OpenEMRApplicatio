package com.cyient.page;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

public class PatientFinderPage {

	private String frameLocator = "fin";
	private By addNewPatientLocator = By.xpath("//button[normalize-space()='Add New Patient']");
	
	private WebDriver driver;

	public PatientFinderPage(WebDriver driver) {
		this.driver = driver;
	}

	public void switchToFinFrame() {
		driver.switchTo().frame(frameLocator);
	}

	public void addNewPatient() {
		driver.findElement(addNewPatientLocator).click();
	}
	
	public void switchToOutOfFinFrame() {
		driver.switchTo().defaultContent();
	}
}
