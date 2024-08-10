package com.html2excel;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import com.html2excel.processor.DataSaver;
import com.html2excel.processor.ExcelCreator;
import com.html2excel.processor.HtmlDataSelector;
import org.apache.poi.ss.usermodel.*;
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

public class Html2Excel extends JFrame {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	static Logger logger = Logger.getLogger(Html2Excel.class);

	private JTextField filePathField;
	private JButton chooseFileButton;
	private JButton goButton;
	private JFileChooser fileChooser;

	public Html2Excel() {
		setTitle("HTML to Excel Converter");
		setSize(400, 150);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setLocationRelativeTo(null);
		setLayout(new FlowLayout());

		ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
		PropertyConfigurator.configure(classLoader.getResource("log4j.properties"));

		filePathField = new JTextField(20);
		filePathField.setEditable(false);
		chooseFileButton = new JButton("Choose File");
		goButton = new JButton("Go");
		fileChooser = new JFileChooser();

		chooseFileButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				int returnValue = fileChooser.showOpenDialog(null);
				if (returnValue == JFileChooser.APPROVE_OPTION) {
					File selectedFile = fileChooser.getSelectedFile();
					filePathField.setText(selectedFile.getAbsolutePath());
				}
			}
		});

		goButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				String pathName = filePathField.getText();
				if (!pathName.isEmpty()) {
					convertHtmlToExcel(pathName);
				} else {
					JOptionPane.showMessageDialog(null, "Please choose a file first!");
				}
			}
		});

		add(filePathField);
		add(chooseFileButton);
		add(goButton);
	}

	private void convertHtmlToExcel(String pathName) {
		HtmlDataSelector htmlDataSelector = new HtmlDataSelector();
		ExcelCreator excelCreator = new ExcelCreator();
		DataSaver dataSaver = new DataSaver();

		try {
			Document doc = htmlDataSelector.htmlParse(pathName);
			Element table = htmlDataSelector.selectTable(doc);
			Workbook workbook = excelCreator.createWorkbook(table);
			dataSaver.saveToExcel(workbook);
			logger.info("Done, saved Excel File on your Desktop");
			JOptionPane.showMessageDialog(null, "Conversion complete! Excel file saved to Desktop.");
		} catch (IOException e) {
			logger.error("An error occurred while converting HTML to Excel", e);
			JOptionPane.showMessageDialog(null, "An error occurred during conversion.");
		}
	}

	public static void main(String[] args) {
		SwingUtilities.invokeLater(() -> {
			Html2Excel app = new Html2Excel();
			app.setVisible(true);
		});
	}
}
