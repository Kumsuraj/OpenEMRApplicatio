package com.cyient.page;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

public class PatientDashboardPage {

	private By patFrameLocator = By.xpath("//iframe[@name='pat']");
	private By medicalRecordDashLocator = By.xpath("//h2[contains(text(),'Medical Record Dashboard')]");
	private WebDriver driver;

	public PatientDashboardPage(WebDriver driver) {
		this.driver = driver;
	}
	
	public void switchTOPatFrame() {
		driver.switchTo().frame(driver.findElement(patFrameLocator));
	}
	
	public String medicalRecordDash() {
		
		String name = new String(
				driver.findElement(medicalRecordDashLocator).getText());
		System.out.println(name);
		return name;
	}

}
