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
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

public class Html2Excel extends JFrame {
	private static final long serialVersionUID = 1L;
	static Logger logger = Logger.getLogger(Html2Excel.class);

	private JTextField filePathField;
	private JButton chooseFileButton;
	private JButton goButton;
	private JFileChooser fileChooser;

	public Html2Excel() {
		setTitle("HTML to Excel Converter");
		setSize(500, 200);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setLocationRelativeTo(null);
		setLayout(new BorderLayout(10, 10));

		ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
		PropertyConfigurator.configure(classLoader.getResource("log4j.properties"));

		// Set the look and feel
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (Exception e) {
			e.printStackTrace();
		}

		// Create the header panel with "Designed For Nixon1993"
		JPanel headerPanel = new JPanel(new BorderLayout());
		JLabel headerLabel = new JLabel("Designed For Nixon1993");
		headerLabel.setBorder(new EmptyBorder(10, 10, 0, 0));

		// Set font to a handwritten style and color to red
		headerLabel.setFont(new Font("Comic Sans MS", Font.BOLD, 16)); // Example of a handwritten-style font
		headerLabel.setForeground(Color.RED); // Set text color to red

		headerPanel.add(headerLabel, BorderLayout.WEST);

		// Create the center panel for file selection and conversion button
		JPanel centerPanel = new JPanel(new FlowLayout());
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

		centerPanel.add(filePathField);
		centerPanel.add(chooseFileButton);
		centerPanel.add(goButton);

		// Create the footer panel with "Created by AlbusPercival91"
		JPanel footerPanel = new JPanel(new BorderLayout());
		JLabel footerLabel = new JLabel("Created by AlbusPercival91");
		footerLabel.setBorder(new EmptyBorder(0, 10, 10, 0));
		footerLabel.setFont(new Font("Arial", Font.ITALIC, 12));
		footerPanel.add(footerLabel, BorderLayout.WEST);

		// Add About menu
		JMenuBar menuBar = new JMenuBar();
		JMenu helpMenu = new JMenu("Help");
		JMenuItem aboutMenuItem = new JMenuItem("About");
		helpMenu.add(aboutMenuItem);
		menuBar.add(helpMenu);
		setJMenuBar(menuBar);

		aboutMenuItem.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				JOptionPane.showMessageDialog(Html2Excel.this,
						"This application is fully compatible with ECDIS Danelec DM700 only.\n May be with DM800, but who knows?",
						"About", JOptionPane.INFORMATION_MESSAGE);
			}
		});

		// Add panels to the frame
		add(headerPanel, BorderLayout.NORTH);
		add(centerPanel, BorderLayout.CENTER);
		add(footerPanel, BorderLayout.SOUTH);
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
