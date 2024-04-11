package com.html2excel.processor;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class ExcelCreator {

	public Workbook createWorkbook(Element table) {
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = createSheet(workbook);
		Row row = createRow(sheet);
		Elements ths = createHeaderRow(table, row);
		createDataRows(table, sheet, ths);
		return workbook;
	}

	private Sheet createSheet(Workbook workbook) {
		Sheet sheet = workbook.createSheet("HTML Table");
		return sheet;
	}

	private Row createRow(Sheet sheet) {
		Row headerRow = sheet.createRow(0);
		return headerRow;
	}

	private Elements createHeaderRow(Element table, Row headerRow) {
		Elements ths = table.select("tr").first().select("th, td").next();

		for (int i = 0; i < ths.size(); i++) {
			Cell headerCell = headerRow.createCell(i);
			headerCell.setCellValue(ths.get(i).text());
		}
		return ths;
	}

	private void createDataRows(Element table, Sheet sheet, Elements ths) {
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
}
