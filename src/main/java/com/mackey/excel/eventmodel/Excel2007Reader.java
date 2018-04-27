package com.mackey.excel.eventmodel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Excel2007Reader extends DefaultHandler implements IReader {
	private SharedStringsTable sst;
	
	private String lastContents;
	private boolean nextlsString;
	
	private boolean isTElement;
	
	private int curRow = 0;//当前行
	private int curCol = 0;//当前列
	private int sheetIndex = -1;
	
	//行数据
	private List<String> rowData = new ArrayList<String>();
	//行数据处理
	private IRowReader rowReader;
	
	private void init(OPCPackage pkg) throws IOException, OpenXML4JException, InvalidFormatException, SAXException {
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();
		XMLReader parser = fetchSheetParser(sst);
		Iterator<InputStream> sheets = r.getSheetsData();
		while(sheets.hasNext())
		{
			curCol = 0;
			sheetIndex ++;
			InputStream sheet = sheets.next();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
		}
	}
	
	public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
		XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
		this.sst = sst;
		parser.setContentHandler(this);
		return parser;
	}
	
	@Override
	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
		if("c".equals(name))
		{
			String cellType = attributes.getValue("t");
			if("s".equals(cellType))
			{
				nextlsString = true;
			} else {
				nextlsString = false;
			}
		}
		
		if("t".equals(name))
		{
			isTElement = true;
		} else {
			isTElement = false;
		}
		
		lastContents = "";
	}
	
	@Override
	public void endElement(String uri, String localName, String name) throws SAXException {
		if(nextlsString)
		{
			int idx = Integer.parseInt(lastContents);
			lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
			nextlsString = false; 
		}
		if(isTElement)
		{
			rowData.add(curCol, lastContents);
			curCol++;
			isTElement = false;
		} else if("v".equals(name)) {
			rowData.add(curCol, lastContents);
			curCol++;
		} else {
			if("row".equals(name))
			{
                if(rowReader == null)
                {
                	throw new IllegalArgumentException("rowReader is not null。。。。");
                }
                
                rowReader.row(sheetIndex, curRow, rowData);
                
				rowData.clear();
				curRow ++;
				curCol = 0;
			}
		}
	}

	@Override
	public void characters(char[] c, int start, int length) throws SAXException {
		lastContents += new String(c,start,length);
	}

	public void process(String path) throws Exception {
		OPCPackage 	pkg = OPCPackage.open(path);
		init(pkg);
	}

	public void process(File file) throws Exception {
		OPCPackage 	pkg = OPCPackage.open(file);
		init(pkg);
	}

	public void process(InputStream in) throws Exception {
		OPCPackage 	pkg = OPCPackage.open(in);
		init(pkg);
	}
	
	public IRowReader getRowReader() {
		return rowReader;
	}

	public void setRowReader(IRowReader rowReader) {
		this.rowReader = rowReader;
	}
}
