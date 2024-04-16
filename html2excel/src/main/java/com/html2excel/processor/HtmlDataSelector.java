package com.html2excel.processor;

import java.io.File;
import java.io.IOException;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

public class HtmlDataSelector {

	public Document htmlParse(String pathName) throws IOException {
		return Jsoup.parse(new File(pathName), "UTF-8");
	}

	public Element selectTable(Document doc) {
		return doc.select("table").first();
	}
}
