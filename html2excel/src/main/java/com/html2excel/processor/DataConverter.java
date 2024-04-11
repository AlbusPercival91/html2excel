package com.html2excel.processor;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;

public class DataConverter {

	public void convertToExcel(Workbook workbook) throws IOException {
		String desktopPath = System.getProperty("user.home") + "/Desktop/";

		try (FileOutputStream fos = new FileOutputStream(desktopPath + "ExcelOutput.xlsx")) {
			workbook.write(fos);
		} finally {
			workbook.close();
		}
	}
}
