package com.gdubina.tool.langutil;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;


public class ToolExport {
	
	private static final String DIR_VALUES = "values"; 
	
	private DocumentBuilder builder;
	private File outExcelFile;
	private String project;
	private Map<String, Integer> keysIndex;
	private PrintStream out;
	private String inputFileName;
	
	public ToolExport(PrintStream out) throws ParserConfigurationException{
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		builder = dbf.newDocumentBuilder();
		this.out = out == null ? System.out : out;
	}

	public static void main(String[] args) throws SAXException, IOException, ParserConfigurationException {
		if(args == null || args.length == 0){
			System.out.println("Project folder doesn't exists");
			return;
		}
		run(null, args[0], args.length > 1 ? args[1] : null, null);
	}
	
	public static void run(String projectDir, String outputFile) throws SAXException, IOException, ParserConfigurationException {
		run(null, projectDir, outputFile, null);
	}

	public static void run(String projectDir, String outputFile, String inFileName) throws SAXException, IOException, ParserConfigurationException {
		run(null, projectDir, outputFile, inFileName);
	}
	
	public static void run(PrintStream out, String projectDir, String outputFile, String inFileName) throws SAXException, IOException, ParserConfigurationException {
		ToolExport tool = new ToolExport(out);
		if(projectDir == null || "".equals(projectDir)){
			tool.out.println("Project folder doesn't exists");
			return;
		}
		File project = new File(projectDir);
		tool.outExcelFile = new File(outputFile != null ? outputFile : "exported_strings_" + System.currentTimeMillis() + ".xls");
		tool.project = project.getName();
		tool.inputFileName = inFileName == null ? "strings.xml" : inFileName;
		tool.export(project);
	}
	
	private void export(File project) throws SAXException, IOException{
		File res = new File(project, "res");
		if (res == null || !res.exists()) {
			System.out.println("res folder doesn't exists");
			return;
		}
		for(File dir : res.listFiles()){
			if(!dir.isDirectory() || !dir.getName().startsWith(DIR_VALUES)){
				continue;
			}
			String dirName = dir.getName();
			if(dirName.equals(DIR_VALUES)){
				keysIndex = exportDefLang(dir);
			}
		}
		if (keysIndex == null) {
			System.out.println("res/values/ folder doesn't exists");
			return;
		}
		for(File dir : res.listFiles()){
			if(!dir.isDirectory() || !dir.getName().startsWith(DIR_VALUES)){
				continue;
			}
			String dirName = dir.getName();
			if(!dirName.equals(DIR_VALUES)){
				int index = dirName.indexOf('-');
				if(index == -1)
					continue;
				String lang = dirName.substring(index + 1);
				exportLang(lang, dir);
			}
		}
	}
	
	private void exportLang(String lang, File valueDir) throws FileNotFoundException, IOException, SAXException{
		File stringFile = new File(valueDir, inputFileName);
		if(!stringFile.exists()){
			return;
		}
		exportLangToExcel(project, lang, getStrings(stringFile), outExcelFile, keysIndex);
	}
	
	private Map<String, Integer> exportDefLang(File valueDir) throws FileNotFoundException, IOException, SAXException{
		File stringFile = new File(valueDir, inputFileName);
		if(!stringFile.exists()){
			return null;
		}
		return exportDefLangToExcel(project, getStrings(stringFile), outExcelFile);
	}
	
	private NodeList getStrings(File f) throws SAXException, IOException{
		Document dom = builder.parse(f);
		return dom.getDocumentElement().getChildNodes();
	}
	
	private static HSSFCellStyle createTilteStyle(HSSFWorkbook wb){
		HSSFFont bold = wb.createFont();
		bold.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		
		HSSFCellStyle style = wb.createCellStyle();
		style.setFont(bold);
		style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style.setWrapText(true);
		
		return style;
	}
	
	private static HSSFCellStyle createCommentStyle(HSSFWorkbook wb){
	
		HSSFFont commentFont = wb.createFont();
		commentFont.setColor(HSSFColor.GREEN.index);
		commentFont.setItalic(true);
		commentFont.setFontHeightInPoints((short)12);
		
		HSSFCellStyle commentStyle = wb.createCellStyle();
		commentStyle.setFont(commentFont);
		return commentStyle;
	}
	
	private static HSSFCellStyle createPlurarStyle(HSSFWorkbook wb){
		
		HSSFFont commentFont = wb.createFont();
		commentFont.setColor(HSSFColor.GREY_50_PERCENT.index);
		commentFont.setItalic(true);
		commentFont.setFontHeightInPoints((short)12);
		
		HSSFCellStyle commentStyle = wb.createCellStyle();
		commentStyle.setFont(commentFont);
		return commentStyle;
	}
	
	private static HSSFCellStyle createKeyStyle(HSSFWorkbook wb){
		HSSFFont bold = wb.createFont();
		bold.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		bold.setFontHeightInPoints((short)11);
		
		HSSFCellStyle keyStyle = wb.createCellStyle();
		keyStyle.setFont(bold);
		
		/*keyStyle.setFillForegroundColor(HSSFColor.LEMON_CHIFFON.index);
		keyStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);*/
		return keyStyle;
	}
	
	private static HSSFCellStyle createTextStyle(HSSFWorkbook wb){
		HSSFFont plain = wb.createFont();
		plain.setFontHeightInPoints((short)12);
	
		HSSFCellStyle textStyle = wb.createCellStyle();
		textStyle.setFont(plain);
		
		return textStyle;
	}
	
	private static HSSFCellStyle createMissedStyle(HSSFWorkbook wb){
	
		HSSFCellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(HSSFColor.RED.index);
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		
		return style;
	}
	
	private static void createTilte(HSSFWorkbook wb, HSSFSheet sheet){
		HSSFRow titleRow = sheet.getRow(0);
		
		HSSFCell cell = titleRow.createCell(0);
		cell.setCellStyle(createTilteStyle(wb));
		cell.setCellValue("KEY");
		
		sheet.setColumnWidth(cell.getColumnIndex(), (40 * 256));
	}
	
	private static void addLang2Tilte(HSSFWorkbook wb, HSSFSheet sheet, String lang){
		HSSFRow titleRow = sheet.getRow(0);
		HSSFCell cell = titleRow.createCell((int)titleRow.getLastCellNum());
		cell.setCellStyle(createTilteStyle(wb));
		cell.setCellValue(lang);
		
		sheet.setColumnWidth(cell.getColumnIndex(), (60 * 256));
	}
	
	
	private Map<String, Integer> exportDefLangToExcel(String project, NodeList strings, File f) throws FileNotFoundException, IOException{
		out.println();
		out.println("Start processing DEFAULT language");
		
		Map<String, Integer> keys = new HashMap<String, Integer>();
		
		HSSFWorkbook wb = new HSSFWorkbook();
		
		HSSFCellStyle commentStyle = createCommentStyle(wb);
		HSSFCellStyle plurarStyle = createPlurarStyle(wb);
		HSSFCellStyle keyStyle = createKeyStyle(wb);
		HSSFCellStyle textStyle = createTextStyle(wb);
	
		HSSFSheet sheet;
		sheet = wb.createSheet(project);
		
		int rowIndex = 0;
		sheet.createRow(rowIndex++);
		createTilte(wb, sheet);
		addLang2Tilte(wb, sheet, "default");
		
		for(int i = 0; i < strings.getLength(); i++){
			Node item = strings.item(i);
			if(item.getNodeType() == Node.TEXT_NODE){
				
			} 
			if(item.getNodeType() == Node.COMMENT_NODE){
				HSSFRow row = sheet.createRow(rowIndex++);
				HSSFCell cell = row.createCell(0);
				cell.setCellValue(String.format("/** %s **/", item.getTextContent()));
				cell.setCellStyle(commentStyle);
				
				sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 0, 255));	
			}
			if("string".equals(item.getNodeName())){
				Node translatable = item.getAttributes().getNamedItem("translatable");
				if(translatable != null && "false".equals(translatable.getNodeValue())){
					continue;
				}
				String key = item.getAttributes().getNamedItem("name").getNodeValue();
				keys.put(key, rowIndex);
				
				HSSFRow row = sheet.createRow(rowIndex++);
				
				HSSFCell cell = row.createCell(0);
				cell.setCellValue(key);
				cell.setCellStyle(keyStyle);
				
				cell = row.createCell(1);
				cell.setCellStyle(textStyle);
				cell.setCellValue(item.getTextContent());
			} else if("plurals".equals(item.getNodeName())){
				String key = item.getAttributes().getNamedItem("name").getNodeValue();
				String plurarName = key; 
				
				HSSFRow row = sheet.createRow(rowIndex++);
				HSSFCell cell = row.createCell(0);
				cell.setCellValue(String.format("//plurals: %s", plurarName));
				cell.setCellStyle(plurarStyle);
				
				NodeList items = item.getChildNodes();
				for(int j = 0; j < items.getLength(); j++){
					Node plurarItem = items.item(j);
					if("item".equals(plurarItem.getNodeName())){
						String itemKey = plurarName + "#" + plurarItem.getAttributes().getNamedItem("quantity").getNodeValue();
						keys.put(itemKey, rowIndex);
						
						HSSFRow itemRow = sheet.createRow(rowIndex++);
						
						HSSFCell itemCell = itemRow.createCell(0);
						itemCell.setCellValue(itemKey);
						itemCell.setCellStyle(keyStyle);
						
						itemCell = itemRow.createCell(1);
						itemCell.setCellStyle(textStyle);
						itemCell.setCellValue(plurarItem.getTextContent());
					}
				}
			} else if("string-array".equals(item.getNodeName())){
				String key = item.getAttributes().getNamedItem("name").getNodeValue();
				String arrayName = key;

				HSSFRow row = sheet.createRow(rowIndex++);
				HSSFCell cell = row.createCell(0);
				cell.setCellValue(String.format("//string-array: %s", arrayName));
				cell.setCellStyle(plurarStyle);

				NodeList items = item.getChildNodes();
				int idx = 0;
				for(int j = 0; j < items.getLength(); j++){
					Node arrayItem = items.item(j);
					if("item".equals(arrayItem.getNodeName())){
						String itemKey = arrayName + "." + idx++;
						keys.put(itemKey, rowIndex);

						HSSFRow itemRow = sheet.createRow(rowIndex++);

						HSSFCell itemCell = itemRow.createCell(0);
						itemCell.setCellValue(itemKey);
						itemCell.setCellStyle(keyStyle);

						itemCell = itemRow.createCell(1);
						itemCell.setCellStyle(textStyle);
						itemCell.setCellValue(arrayItem.getTextContent());
					}
				}
			}
		}
		sheet.createFreezePane(1, 1);
		
		FileOutputStream outFile = new FileOutputStream(f);
		wb.write(outFile);
		outFile.close();
		
		out.println("DEFAULT language was processed");
		return keys;
	}
	
	private void exportLangToExcel(String project, String lang, NodeList strings, File f, Map<String, Integer> keysIndex) throws FileNotFoundException, IOException{
		out.println();
		out.println(String.format("Start processing: '%s'", lang));
		Set<String> missedKeys = new HashSet<String>(keysIndex.keySet());
		
		HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(f));
		
		HSSFCellStyle textStyle = createTextStyle(wb);
		
		HSSFSheet sheet = wb.getSheet(project);
		addLang2Tilte(wb, sheet, lang);
		for(int i = 0; i < strings.getLength(); i++){
			Node item = strings.item(i);

			if("string".equals(item.getNodeName())){
				Node translatable = item.getAttributes().getNamedItem("translatable");
				if(translatable != null && "false".equals(translatable.getNodeValue())){
					continue;
				}
				String key = item.getAttributes().getNamedItem("name").getNodeValue();
				Integer index = keysIndex.get(key);
				if(index == null){
					out.println("\t" + key + " - row does not exist");
					continue;
				}
				
				missedKeys.remove(key);
				HSSFRow row = sheet.getRow(index);
				
				HSSFCell cell = row.createCell((int)row.getLastCellNum());
				cell.setCellValue(item.getTextContent());
				cell.setCellStyle(textStyle);
			} else if("plurals".equals(item.getNodeName())){
				String key = item.getAttributes().getNamedItem("name").getNodeValue();
				String plurarName = key;
				
				NodeList items = item.getChildNodes();
				for(int j = 0; j < items.getLength(); j++){
					Node plurarItem = items.item(j);
					if("item".equals(plurarItem.getNodeName())){
						key = plurarName + "#" + plurarItem.getAttributes().getNamedItem("quantity").getNodeValue();
						Integer index = keysIndex.get(key);
						if(index == null){
							out.println("\t" + key + " - row does not exist");
							continue;
						}
						missedKeys.remove(key);
						
						HSSFRow row = sheet.getRow(index);
						
						HSSFCell cell = row.createCell((int)row.getLastCellNum());
						cell.setCellValue(plurarItem.getTextContent());
						cell.setCellStyle(textStyle);
					}
				}
			} else if ("string-array".equals(item.getNodeName())) {
				String key = item.getAttributes().getNamedItem("name").getNodeValue();
				String arrayName = key;

				NodeList items = item.getChildNodes();
				int idx = 0;
				for(int j = 0; j < items.getLength(); j++){
					Node arrayItem = items.item(j);
					if("item".equals(arrayItem.getNodeName())){
						key = arrayName + "." + idx++;
						Integer index = keysIndex.get(key);
						if(index == null){
							out.println("\t" + key + " - row does not exist");
							continue;
						}
						missedKeys.remove(key);

						HSSFRow row = sheet.getRow(index);

						HSSFCell cell = row.createCell((int) row.getLastCellNum());
						cell.setCellValue(arrayItem.getTextContent());
						cell.setCellStyle(textStyle);
					}
				}
			}
		}
		
		HSSFCellStyle missedStyle = createMissedStyle(wb);
		
		if(!missedKeys.isEmpty()){
			out.println("  MISSED KEYS:");
		}
		for(String missedKey : missedKeys){
			out.println("\t" + missedKey);
			Integer index = keysIndex.get(missedKey);
			HSSFRow row = sheet.getRow(index);
			HSSFCell cell = row.createCell((int)row.getLastCellNum());
			cell.setCellStyle(missedStyle);
		}
		
		FileOutputStream outStream = new FileOutputStream(f);
		wb.write(outStream);
		outStream.close();
		
		if(missedKeys.isEmpty()){
			out.println(String.format("'%s' was processed", lang));
		}else{
			out.println(String.format("'%s' was processed with MISSED KEYS - %d" , lang, missedKeys.size()));
		}
	}
	
	/*private static String deluminateText(String text){
		return text.replace("\\'", "'").replace("\\\"", "\"");
	}*/
}
