package com.html2excel;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import java.io.*;
import java.util.Scanner;

public class Html2ExcelConverter {
	static Logger logger = Logger.getLogger(Html2ExcelConverter.class);

	public static void main(String[] args) {
		ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
		PropertyConfigurator.configure(classLoader.getResource("log4j.properties"));

		logger.info("Please insert Path to File (Example: C:\\Users\\Serge\\Desktop\\08042024_012308.html" + "\n");

		try (Scanner scan = new Scanner(System.in)) {
			String pathName = scan.nextLine();

			Document doc = htmlParse(pathName);

			Element table = selectTable(doc);

			Workbook workbook = createWorkbook(table);

			convertToExcel(workbook);
		} catch (IOException e) {
			logger.error("An error occurred while converting HTML to Excel", e);
		}

		logger.info("Done, saved Excel File in your Desktop");
	}

	private static Document htmlParse(String pathName) throws IOException {
		Document doc = Jsoup.parse(new File(pathName), "UTF-8");
		return doc;
	}

	private static Element selectTable(Document doc) {
		Element table = doc.select("table").first();
		return table;
	}

	private static Workbook createWorkbook(Element table) {
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = createSheet(workbook);
		Row row = createRow(sheet);
		Elements ths = createHeaderRow(table, row);
		createDataRows(table, sheet, ths);
		return workbook;
	}

	private static Sheet createSheet(Workbook workbook) {
		Sheet sheet = workbook.createSheet("HTML Table");
		return sheet;
	}

	private static Row createRow(Sheet sheet) {
		Row headerRow = sheet.createRow(0);
		return headerRow;
	}

	private static Elements createHeaderRow(Element table, Row headerRow) {
		Elements ths = table.select("tr").first().select("th, td").next();

		for (int i = 0; i < ths.size(); i++) {
			Cell headerCell = headerRow.createCell(i);
			headerCell.setCellValue(ths.get(i).text());
		}
		return ths;
	}

	private static void createDataRows(Element table, Sheet sheet, Elements ths) {
		Elements rows = table.select("tr");

		for (int i = 1; i < rows.size(); i++) {
			Row row = sheet.createRow(i);
			Elements cols = rows.get(i).select("td");

			for (int j = 0; j < cols.size(); j++) {
				Cell cell = row.createCell(j);
				String cellValue = cols.get(j).text();

				// Check if the column is not 'Leg distance (NM)'
				if (!ths.get(j).text().equals("Leg distance (NM)")) {
					cellValue = cellValue.replace("�", "\u00B0" + ' '); // replace � with degree symbol
				} else {
					cellValue = cellValue.replace("�", ""); // remove �
				}

				cell.setCellValue(cellValue);
			}
		}
	}

	private static void convertToExcel(Workbook workbook) throws IOException {
		String desktopPath = System.getProperty("user.home") + "/Desktop/";

		try (FileOutputStream fos = new FileOutputStream(desktopPath + "ExcelOutput.xlsx")) {
			workbook.write(fos);
		} finally {
			workbook.close();
		}
	}

}
