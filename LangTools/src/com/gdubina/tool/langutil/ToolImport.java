package com.gdubina.tool.langutil;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintStream;
import java.util.Iterator;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class ToolImport {

	private DocumentBuilder builder;
	private File outResDir;
	private PrintStream out;

	public ToolImport(PrintStream out) throws ParserConfigurationException {
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		builder = dbf.newDocumentBuilder();
		this.out = out == null ? System.out : out;
	}

	public static void main(String[] args) throws FileNotFoundException, IOException, ParserConfigurationException, TransformerException {
		if(args == null || args.length == 0){
			System.out.println("File name is missed");
			return;
		}
		run(args[0]);
	}
	
	public static void run(String input) throws FileNotFoundException, IOException, ParserConfigurationException, TransformerException{
		run(input, null);
	}
	
	public static void run(String input, String stringsFile) throws FileNotFoundException, IOException, ParserConfigurationException, TransformerException{
		if(input == null || "".equals(input)){
			System.out.println("File name is missed");
			return;
		}
		HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(new File(input)));
		HSSFSheet sheet = wb.getSheetAt(0);
		

		ToolImport tool = new ToolImport(null);
		tool.outResDir = new File("out/" + sheet.getSheetName()+ "/res");
		tool.outResDir.mkdirs();
		tool.parse(sheet, stringsFile);		
	}
	
	public static void run(PrintStream out, String projectDir, String input, String stringsFile) throws FileNotFoundException, IOException, ParserConfigurationException, TransformerException{
		ToolImport tool = new ToolImport(out);
		if(input == null || "".equals(input)){
			tool.out.println("File name is missed");
			return;
		}
		
		HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(new File(input)));
		HSSFSheet sheet = wb.getSheetAt(0);
		

		tool.outResDir = new File(projectDir, "/res");
		//tool.outResDir.mkdirs();
		tool.parse(sheet, stringsFile);		
	}

	private void parse(HSSFSheet sheet, String stringsFile) throws IOException, TransformerException {
		Row row = sheet.getRow(0);
		Iterator<Cell> cells = row.cellIterator();
		cells.next();// ignore key
		int i = 1;
		while (cells.hasNext()) {
			String lang = cells.next().toString();
			generateLang(sheet, lang, i, stringsFile);
			i++;
		}
	}

	private void generateLang(HSSFSheet sheet, String lang, int column, String stringsFile) throws IOException, TransformerException {
		
		Document dom = builder.newDocument();
		Element root = dom.createElement("resources");
		dom.appendChild(root);
		
		Iterator<Row> iterator = sheet.rowIterator();
		iterator.next();//ignore first row;
		Element pluralsNode = null;
		Element arraysNode = null;
		String plurarName = null;
		String arrayName = null;

		while (iterator.hasNext()) {
			HSSFRow row = (HSSFRow) iterator.next();
			Cell cell = row.getCell(0);// android key
			if (cell == null) {
				continue;
			}
			String key = cell.toString();
			if (key == null || "".equals(key)){
				root.appendChild(dom.createTextNode(""));
				continue;
			}
			if(key.startsWith("/**")){
				root.appendChild(dom.createComment(key.substring(3, key.length() - 3)));
				continue;
			}
			
			if(key.startsWith("//")){
				root.appendChild(dom.createComment(key.substring(2)));
				continue;
			}
			
			int plurarIndex = key.indexOf("#");
			int arrayDotIndex = key.indexOf(".");
			if(plurarIndex == -1 && arrayDotIndex == -1){//string
				Cell valueCell = row.getCell(column);
				if(valueCell == null){
					addEmptyKeyValue(dom, root, key);
					continue;
				}
				String value = valueCell.toString();// value
				
				if(value.isEmpty()){
					addEmptyKeyValue(dom, root, key);
				}else {
					Element node = dom.createElement("string");
					node.setAttribute("name", key);
					node.setTextContent(value);
					root.appendChild(node);
				}
			} else if (arrayDotIndex != -1) {
				Cell valueCell = row.getCell(column);
				String value = "";
				if(valueCell != null){
					value = valueCell.toString();// value
				}
				String arrayNameNew = key.substring(0, arrayDotIndex);
				if(!arrayNameNew.equals(arrayName)){
					arrayName = arrayNameNew;
					arraysNode = dom.createElement("string-array");
					arraysNode.setAttribute("name", arrayName);
				}
				Element item = dom.createElement("item");
				item.setTextContent(value);

				arraysNode.appendChild(item);

				root.appendChild(arraysNode);
			}else{
				Cell valueCell = row.getCell(column);
				String value = "";
				if(valueCell != null){
					value = valueCell.toString();// value
				}
				String plurarNameNew = key.substring(0, plurarIndex);
				String quantity = key.substring(plurarIndex + 1);
				if(!plurarNameNew.equals(plurarName)){
					plurarName = plurarNameNew; 
					pluralsNode = dom.createElement("plurals");
					pluralsNode.setAttribute("name", plurarName);
				}
				Element item = dom.createElement("item");
				item.setAttribute("quantity", quantity);
				item.setTextContent(value);
				
				pluralsNode.appendChild(item);
				
				root.appendChild(pluralsNode);				
			}

		}
		
		save(dom, lang, stringsFile);
	}
	
	private static void addEmptyKeyValue(Document dom, Element root, String key){
		root.appendChild(dom.createComment(String.format(" TODO: string name=\"%s\" ", key)));
	} 

	private void save(Document doc, String lang, String fileName) throws TransformerException {
		File dir;
		if("default".equals(lang) || lang == null || "".equals(lang)){
			dir = new File(outResDir, "values");
		}else{
			dir = new File(outResDir, "values-" + lang);
		}
		dir.mkdir();
		
		//DOMUtils.prettyPrint(doc);
		TransformerFactory transformerFactory = TransformerFactory.newInstance();
		Transformer transformer = transformerFactory.newTransformer();
		transformer.setOutputProperty(OutputKeys.INDENT, "yes");
		transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
		
		DOMSource source = new DOMSource(doc);
		if (fileName == null || fileName.length() == 0) {
			fileName = "strings.xml";
		}
		StreamResult result = new StreamResult(new File(dir, fileName));

		transformer.transform(source, result);
	}
	
/*
	private static String eluminateText(String text) {
		return text.replace("'", "\\'").replace("\"", "\\\"");
	}
*/
}
