package com.html2excel;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.ss.usermodel.*;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import com.html2excel.processor.HrmlDataSelector;
import com.html2excel.processor.ExcelCreator;
import com.html2excel.processor.DataSaver;
import java.io.*;
import java.util.Scanner;

public class Html2Excel {
	static Logger logger = Logger.getLogger(Html2Excel.class);

	public static void main(String[] args) {
		ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
		PropertyConfigurator.configure(classLoader.getResource("log4j.properties"));

		logger.info("Please insert Path to File (Example: C:\\Users\\Serge\\Desktop\\08042024_012308.html" + "\n");

		HrmlDataSelector htmlDataSelector = new HrmlDataSelector();
		ExcelCreator excelCreator = new ExcelCreator();
		DataSaver dataSaver = new DataSaver();

		try (Scanner scan = new Scanner(System.in)) {
			String pathName = scan.nextLine();

			Document doc = htmlDataSelector.htmlParse(pathName);

			Element table = htmlDataSelector.selectTable(doc);

			Workbook workbook = excelCreator.createWorkbook(table);

			dataSaver.saveToExcel(workbook);
		} catch (IOException e) {
			logger.error("An error occurred while converting HTML to Excel", e);
		}

		logger.info("Done, saved Excel File in your Desktop");
	}

}
