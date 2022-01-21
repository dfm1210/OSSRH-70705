package cc.ebatis.create;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import cc.ebatis.annotation.EnableExcelMaker;
import cc.ebatis.annotation.ExcelField;
import cc.ebatis.annotation.ExcelTitle;
import cc.ebatis.exception.NoEnableExcelMakerException;
import cc.ebatis.util.ConvertUtil;

public class CreateExcel<T> {
	
	private HSSFWorkbook info = null;
	private HSSFSheet sheet = null;
	private HSSFCellStyle cellStyle = null;
	
	public HSSFWorkbook getHSSFWorkbook() {
		return info;
	}
	
	public HSSFSheet getHSSFSheet() {
		return sheet;
	}
	
	public void write(File file) {
		try {
			info.write(file);
			info.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public void create(List<T> list, String sheetName) throws NoEnableExcelMakerException {
		
		Class<? extends Object> class1 = list.get(0).getClass();
		EnableExcelMaker enableExcelMaker = class1.getAnnotation(EnableExcelMaker.class);
		ExcelTitle excelTitle;
		String title = null;
		Field[] fields;
		Map<Integer,String[]> map = new HashMap<Integer,String[]>();
		Set<Integer> keySet;
		List<Integer> mergeNums = new ArrayList<Integer>();
		
		if(enableExcelMaker == null) {
			throw new NoEnableExcelMakerException("Can't find @EnableExcelMaker annotation on class");
		}
		excelTitle = class1.getAnnotation(ExcelTitle.class);
		
		if(excelTitle != null) {
			title = excelTitle.value();
		}
		
		fields = class1.getDeclaredFields();
		for(Field x : fields) {
			ExcelField excelField = x.getAnnotation(ExcelField.class);
			if(excelField == null) {
				continue;
			}
			String name = excelField.name();
			int position = excelField.position();
			String width = String.valueOf(excelField.width());
			String merge = String.valueOf(excelField.merge());
			String fieldName = x.getName();
			String[] infos = new String[]{name, fieldName, width, merge};
			map.put(position, infos);
			
			if(Boolean.valueOf(merge)) {
				mergeNums.add(position);
			}
			
		}

		
		info = new HSSFWorkbook();
		sheet = info.createSheet(sheetName);
		cellStyle = info.createCellStyle();
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		

		int firstLineIndex = 0;
		
		if(title != null) {
			firstLineIndex = 1;
		}
		
		HSSFRow createRow = sheet.createRow(firstLineIndex);
		
		HSSFCellStyle crs = getStyleBold(info);
		HSSFCellStyle cellCrs = getStyle(info);
		
		keySet = map.keySet();
		int maxCell = -1;
		for(Integer x : keySet) {
			if(x > maxCell) {
				maxCell = x;
			}
			String string = map.get(x)[0];
			//sheet.setDefaultColumnStyle(x, cellStyle);
			HSSFCell createCell = createRow.createCell(x,CellType.STRING);
			createCell.setCellStyle(crs);
			createCell.setCellValue(string);
			String lengthStr = map.get(x)[2];
			if(!lengthStr.equals("-1")) {
				sheet.setColumnWidth(x, Integer.parseInt(map.get(x)[2]) * 400);
			}
			
		}
		
		if(firstLineIndex == 1) {
			HSSFRow titleRow = sheet.createRow(0);
			titleRow.setHeight((short)666);
			HSSFCell createCell = titleRow.createCell(0);
			createCell.setCellType(CellType.STRING);
			createCell.setCellValue(title);
	        CellRangeAddress cra =new CellRangeAddress(0, 0, 0, maxCell);  
	        sheet.addMergedRegion(cra);
	        createCell.setCellStyle(getStyleTitle(info));
		}
		
		Map<Integer,String> beforeMap = new HashMap<Integer,String>();
		Map<Integer,Integer> beforeMapMerge = new HashMap<Integer,Integer>();
		
		for(Integer x:mergeNums) {
			beforeMap.put(x, null);
			beforeMapMerge.put(x, 0);
		}
				
		
		for(int i = 0; i < list.size(); i++) {
			T t = list.get(i);
			Class<? extends Object> class2 = t.getClass();
			HSSFRow row = sheet.createRow(i + 1 + firstLineIndex);
			for(Integer x : keySet) {
				String string = map.get(x)[1];
				String method = "get" + ConvertUtil.upperCase(string);
				Object invoke = null;
				try {
					invoke = class2.getMethod(method).invoke(t);
				} catch (Exception e) {
					e.printStackTrace();
				}
				if(invoke == null) {
					continue;
				}
				
				String typeName = invoke.getClass().getTypeName();
				HSSFCell createCell = row.createCell(x);
				createCell.setCellStyle(cellCrs);
				switch(typeName) {
				case "java.lang.String":
					createCell.setCellType(CellType.STRING);
					createCell.setCellValue((String)invoke);
					break;
				case "java.lang.Double":
					createCell.setCellType(CellType.NUMERIC);
					createCell.setCellValue((Double)invoke);
					break;
				case "java.lang.Short":
					createCell.setCellType(CellType.NUMERIC);
					createCell.setCellValue((Short)invoke);
					break;
				case "java.lang.Long":
					createCell.setCellType(CellType.NUMERIC);
					createCell.setCellValue((Long)invoke);
					break;
				case "java.lang.Integer":
					createCell.setCellType(CellType.NUMERIC);
					createCell.setCellValue((Integer)invoke);
					break;
				case "java.lang.Boolean":
					createCell.setCellType(CellType.BOOLEAN);
					createCell.setCellValue((Boolean)invoke);
					break;
				case "java.util.Date":
					Date date = (Date)invoke;
					SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
					String format2 = format.format(date);
					createCell.setCellType(CellType.STRING);
					createCell.setCellValue(format2);
					break;
				}
				
				
				Set<Integer> keySet2 = beforeMap.keySet();
				
				if(!keySet2.contains(x)) {
					continue;
				}
				
				String string2 = beforeMap.get(x);
					
					String invokStr = String.valueOf(invoke);
					if(string2 != null && string2.equals(invokStr)) {
						
						if(i == list.size() - 1) {
							sheet.addMergedRegion(new CellRangeAddress(i + firstLineIndex - beforeMapMerge.get(x),
									i + 1 + firstLineIndex,x,x));
						}
						
						Integer thesIndex = beforeMapMerge.get(x);
						beforeMapMerge.put(x, thesIndex + 1);
						
					}else {
						if(beforeMapMerge.get(x) != null && beforeMapMerge.get(x) != 0) {
							sheet.addMergedRegion(new CellRangeAddress(i + firstLineIndex - beforeMapMerge.get(x),
									i + firstLineIndex,x,x));
						}
						beforeMapMerge.put(x, 0);
					}
					beforeMap.put(x, invokStr);
				
				
				
			}
		}
		
	}
	
	/**
	 * Get headline style
	 * @param info
	 * @return
	 */
	private HSSFCellStyle getStyleTitle(HSSFWorkbook info) {
		HSSFCellStyle style = getStyle(info);
		HSSFFont crsFont = info.createFont();
		crsFont.setBold(true);
		crsFont.setFontHeightInPoints((short) 14);//设置字体大小    
		style.setFont(crsFont);
		return style;
	}
	
	/**
	 * Get header style
	 * @param info
	 * @return
	 */
	private HSSFCellStyle getStyleBold(HSSFWorkbook info) {
		HSSFCellStyle style = getStyle(info);
		HSSFFont crsFont = info.createFont();
		crsFont.setBold(true);
		style.setFont(crsFont);
		return style;
	}
	
	/**
	 * Get normal cell style (center, border)
	 * @param info
	 * @return
	 */
	private HSSFCellStyle getStyle(HSSFWorkbook info) {
		HSSFCellStyle crs = info.createCellStyle();
		crs.setWrapText(true);
		crs.setBorderBottom(BorderStyle.THIN);
		crs.setBorderLeft(BorderStyle.THIN);    
		crs.setBorderTop(BorderStyle.THIN);   
		crs.setBorderRight(BorderStyle.THIN);
		crs.setAlignment(HorizontalAlignment.CENTER);
		crs.setVerticalAlignment(VerticalAlignment.CENTER);
		return crs;
	}
	
}
