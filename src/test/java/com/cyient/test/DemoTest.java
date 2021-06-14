package com.cyient.test;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DemoTest {

	public static void main(String[] args) throws IOException {

		FileInputStream file = new FileInputStream("src/test/resources/testdata/openEMRData.xlsx");

		XSSFWorkbook book = new XSSFWorkbook(file);

		XSSFSheet sheet = book.getSheet("validCredentialTest");

		int rowCount = sheet.getPhysicalNumberOfRows();
		System.out.println(rowCount);
		int cellCount = sheet.getRow(0).getPhysicalNumberOfCells();
		System.out.println(cellCount);
		/*
		 * String cellValue = cell.getStringCellValue(); System.out.println(cellValue);
		 */

		Object[][] main = new Object[rowCount-1][cellCount];

		for (int i = 1; i < rowCount; i++) {
			for (int j = 0; j < cellCount; j++) {
				XSSFRow row = sheet.getRow(i);
				XSSFCell cell = row.getCell(j);
				DataFormatter format = new DataFormatter();
				String cellValue = format.formatCellValue(cell);
				// System.out.println(cellValue);

				main[i - 1][j] = cellValue;
				System.out.println(main[i][j]);
			}
		}
		/*
		 * System.out.println(main[0][0]); System.out.println(main[1][2]);
		 */
	}

}
