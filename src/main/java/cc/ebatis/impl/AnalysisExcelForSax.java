package cc.ebatis.impl;

import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import cc.ebatis.api.DataHandleAction;
import cc.ebatis.pojo.ActionContext;
import cc.ebatis.pojo.SheetInfo;
import cc.ebatis.util.ReflexObject;

/**
 * Parsing excel using Sax (for bigdata)
 * @author Steve
 *
 * @param <T>
 */
public class AnalysisExcelForSax<T> implements DataHandleAction<T>{

	@Override
	public void prepare(ActionContext<T> act) {

		if(!act.getUseSax()) {
			commit(act);
			return;
		}
		
		try {
			ReflexVO<T> reflexVO = new ReflexVO<T>();
			reflexVO.setAct(act);
			if(act.getFile() == null) {
				processAllSheets(act.getInputStream(), reflexVO);
			}else {
				processAllSheets(act.getFile(), reflexVO);
			}

		} catch (Exception e) {
			rollback(act);
			e.printStackTrace();
		}
		
	}

	@Override
	public boolean commit(ActionContext<T> act) {
		
		act.setResult(true);
		
		return true;
	}

	@Override
	public boolean rollback(ActionContext<T> act) {

		act.setResult(false);
		
		return false;
	}

	private static StylesTable stylesTable;
	
	public void processAllSheets(File file, ReflexVO<T> reflexVO) throws Exception {
		OPCPackage pkg = OPCPackage.open(file,PackageAccess.READ);
		XSSFReader r = new XSSFReader(pkg);
		
		stylesTable = r.getStylesTable();
		SharedStringsTable sst = r.getSharedStringsTable();
		
		XMLReader parser = fetchSheetParser(sst,reflexVO);
		
		
		XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) r.getSheetsData();
		
		Integer sheetSize = 0;
		
		while(sheets.hasNext()) {
			
			InputStream sheet = sheets.next();
			SheetInfo<T> sheetInfo = new SheetInfo<T>();
			sheetInfo.setSheetName(sheets.getSheetName());
			reflexVO.setSheetInfo(sheetInfo);
			sheetSize++;
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
			sheetInfo.setColumn(reflexVO.getListHeader().size());
			sheetInfo.setBlankLineSize(sheetInfo.getBlankLine().size());
			sheetInfo.setErrorLineSize(sheetInfo.getErrorLine().size());
			sheetInfo.setRepeatLineSize(sheetInfo.getRepeatLine().size());
			reflexVO.getAct().addSheets(reflexVO.getSheetInfo());
		}
		reflexVO.getAct().setSheetSize(sheetSize);
		reflexVO.getAct().setResult(true);
	}

	public void processAllSheets(InputStream file, ReflexVO<T> reflexVO) throws Exception {
		OPCPackage pkg = OPCPackage.open(file);
		XSSFReader r = new XSSFReader(pkg);
		IOUtils.closeQuietly(file);
		stylesTable = r.getStylesTable();
		SharedStringsTable sst = r.getSharedStringsTable();
		
		XMLReader parser = fetchSheetParser(sst,reflexVO);
		
		
		XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) r.getSheetsData();
		
		Integer sheetSize = 0;
		
		while(sheets.hasNext()) {
			
			InputStream sheet = sheets.next();
			SheetInfo<T> sheetInfo = new SheetInfo<T>();
			sheetInfo.setSheetName(sheets.getSheetName());
			reflexVO.setSheetInfo(sheetInfo);
			sheetSize++;
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
			sheetInfo.setColumn(reflexVO.getListHeader().size());
			sheetInfo.setBlankLineSize(sheetInfo.getBlankLine().size());
			sheetInfo.setErrorLineSize(sheetInfo.getErrorLine().size());
			sheetInfo.setRepeatLineSize(sheetInfo.getRepeatLine().size());
			reflexVO.getAct().addSheets(reflexVO.getSheetInfo());
		}
		reflexVO.getAct().setSheetSize(sheetSize);
		reflexVO.getAct().setResult(true);
	}
	
	public XMLReader fetchSheetParser(SharedStringsTable sst, ReflexVO<T> reflexVO) throws SAXException {
		
		XMLReader parser =
			XMLReaderFactory.createXMLReader(
					"org.apache.xerces.parsers.SAXParser"
			);
		ContentHandler handler = new SheetHandler<T>(sst, reflexVO);
		parser.setContentHandler(handler);
		return parser;
	}

	/**
	 * An object that encapsulates each piece of recorded information
	 * @author Steve
	 *
	 */
	private static class ReflexVO<T>{
		
		private List<String> listHeader;
		private List<String> rowInfo;
		private Integer lineNum;
		private ActionContext<T> act;
		private SheetInfo<T> sheetInfo;
		private Set<T> distinctSet;
		
		public ReflexVO() {
			this.listHeader = new ArrayList<String>();
			this.rowInfo = new ArrayList<String>();
			this.lineNum = 0;
			this.distinctSet = new HashSet<T>();
		}
		
		public Set<T> getDistinctSet() {
			return distinctSet;
		}

		public SheetInfo<T> getSheetInfo() {
			return sheetInfo;
		}
		public void setSheetInfo(SheetInfo<T> sheetInfo) {
			this.sheetInfo = sheetInfo;
		}
		public ActionContext<T> getAct() {
			return act;
		}
		public void setAct(ActionContext<T> act) {
			this.act = act;
		}
		public List<String> getListHeader() {
			return listHeader;
		}
		public void setListHeader(List<String> listHeader) {
			this.listHeader = listHeader;
		}
		public List<String> getRowInfo() {
			return rowInfo;
		}
		public void setRowInfo(List<String> rowInfo) {
			this.rowInfo = rowInfo;
		}
		public Integer getLineNum() {
			return lineNum;
		}
		public void setLineNum(Integer lineNum) {
			this.lineNum = lineNum;
		}
	}
	
	
	/** 
	 * See org.xml.sax.helpers.DefaultHandler javadocs 
	 */
	private static class SheetHandler<T> extends DefaultHandler {
		private SharedStringsTable sst;
		private String lastContents;
		private boolean nextIsString;
		private ReflexVO<T> reflexVO;
		private boolean distinct;
		private Integer index;
		private boolean isFirstLine = true;

		ReflexObject<T> reflexObject = new ReflexObject<T>();
		
		private SheetHandler(SharedStringsTable sst, ReflexVO<T> reflexVO) {
			this.sst = sst;
			this.reflexVO = reflexVO;
			this.distinct = reflexVO.getAct().getDistinct();
		}
		
		short formatIndex;
        String formatString;
		String cellStyleStr;
		
		public void startElement(String uri, String localName, String name,
				Attributes attributes) throws SAXException{		
			
			if(name.equals("row")) {
				int thisLine = Integer.parseInt(attributes.getValue("r"));
				reflexVO.getSheetInfo().setLine(thisLine - 1);
				int range = thisLine - reflexVO.getLineNum() - 1;	
				if(range > 0) {
					for(int i = 0; i < range; i++) {
						reflexVO.getSheetInfo().addBlankLine(reflexVO.getLineNum() + i + 1);
					}
				}
				
				reflexVO.setLineNum(thisLine);
				if(thisLine == 1) {
					isFirstLine = false;
					reflexVO.setListHeader(new ArrayList<String>());
				}else {
					if(isFirstLine) {
						throw new SAXException("No table head!");
					}
					reflexVO.setRowInfo(new ArrayList<String>());
				}
			}

			if(name.equals("c")) {
				
				index = nameToColumn(attributes.getValue("r"));
				formatIndex = -1;
		        formatString = null;
				cellStyleStr = attributes.getValue("s");
				if (cellStyleStr != null) {
					int styleIndex = Integer.parseInt(cellStyleStr);
		            XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
		            if(style != null) {
		            	formatIndex = style.getDataFormat();
			            formatString = style.getDataFormatString();
		            }
		            if (formatIndex == 14 || formatIndex == 31 || formatIndex == 57 || formatIndex == 176 || formatIndex == 178 || formatIndex == 166) {
		                formatString = "yyyy-MM-dd";
		            }else if(formatIndex == 58 || formatIndex == 177){
		            	formatString = "MM-dd";
		            }else if(formatIndex == 179){
		            	formatString = "yyyy-MM";
		            }else {
		            	formatString = null;
		            }
		            
		        }
				
				String cellType = attributes.getValue("t");
				if(cellType != null && cellType.equals("s")) {
					nextIsString = true;
				} else {
					nextIsString = false;
				}
			}
			// Clear contents cache
			
			lastContents = "";
		}
		
		public void endElement(String uri, String localName, String name)
				throws SAXException {
			if(nextIsString) {
				int idx = Integer.parseInt(lastContents);
				lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
				nextIsString = false;
			}
			
			if(name.equals("v")) {
				if(reflexVO.getLineNum() == 1) {
					List<String> lh = reflexVO.getListHeader();
					if(lh.size() == index) {
						lh.add(lastContents);
					}else {
						lh.add("");
						lh.add(lastContents);
					}
				}else {
					DataFormatter formatter = new DataFormatter();
		            if (formatString != null) {
		            	try {
			            	double parseDouble = Double.parseDouble(lastContents);
			            	String thisStr = formatter.formatRawCellContents(parseDouble, formatIndex, formatString).trim();
			                lastContents = thisStr;
		            	}catch(NumberFormatException e) {
		            	}
		                
		            }

					List<String> ri = reflexVO.getRowInfo();
					
					if(ri.size() == index) {
						ri.add(lastContents);
					}else {
						int allTime = index-ri.size();
						for(int i=0;i < allTime;i++) {
							ri.add("");
						}
						ri.add(lastContents);
					}
				}
			}
			
			if(name.equals("row") && reflexVO.getLineNum() != 1) {
				List<String> listHeader = reflexVO.getListHeader();
				List<String> rowInfo = reflexVO.getRowInfo();
				boolean isAllEmpty = true;
				for(String x : rowInfo) {
					if(x != null && !x.trim().equals("")) {
						isAllEmpty = false;
						break;
					}
				}
				if(isAllEmpty) {
					index = 0;
					SheetInfo<T> sheetInfo = reflexVO.getSheetInfo();
					sheetInfo.addBlankLine(reflexVO.getLineNum());
					return;
				}
				
				int range = (rowInfo.size() - listHeader.size()) * -1;
				if(range > 0) {
					for(int i = 0; i < range; i++) {
						rowInfo.add("");
					}
				}
				T obj = (T)reflexObject.getReflexObject(reflexVO.getAct().getObjects(), 
						listHeader, 
						rowInfo, 
						reflexVO.getSheetInfo().getSheetName(), 
						reflexVO.getLineNum());
				
				if(obj != null) {
					boolean flag = true;
					if(distinct) {
						boolean add = reflexVO.getDistinctSet().add(obj);
						if(!add) {
							reflexVO.getSheetInfo().addRepeatLine(reflexVO.getLineNum());
							flag = false;
						}
					}
					
					if(flag) {
						reflexVO.getSheetInfo().getInfo().add(obj);
						Integer correctLine = reflexVO.getSheetInfo().getCorrectLine() + 1;
						reflexVO.getSheetInfo().setCorrectLine(correctLine);
					}
					
				}
				
				if(obj == null) {
					reflexVO.getSheetInfo().addErrorLine(reflexVO.getLineNum());
				}
				
				index = 0;
				
			}
			
		}

		public void characters(char[] ch, int start, int length)
				throws SAXException {
			lastContents += new String(ch, start, length);
		}
	}
	
	private static int nameToColumn(String name) {
		name = name.replaceAll("\\d+","");
	    int column = -1;    
	    for (int i = 0; i < name.length(); ++i) {    
	        int c = name.charAt(i);    
	        column = (column + 1) * 26 + c - 'A';
	    }    
	    return column;    
	}

}