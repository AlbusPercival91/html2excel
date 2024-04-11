package com.html2excel.processor;

import java.io.File;
import java.io.IOException;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

public class HrmlDataSelector {

	public Document htmlParse(String pathName) throws IOException {
		Document doc = Jsoup.parse(new File(pathName), "UTF-8");
		return doc;
	}

	public Element selectTable(Document doc) {
		Element table = doc.select("table").first();
		return table;
	}
}
