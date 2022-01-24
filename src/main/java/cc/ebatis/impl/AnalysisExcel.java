package cc.ebatis.impl;

import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cc.ebatis.api.DataHandleAction;
import cc.ebatis.pojo.ActionContext;
import cc.ebatis.pojo.SheetInfo;
import cc.ebatis.util.ReflexObject;

/**
 * Analyzing the contents of Excel
 * @author Steve
 *
 */
public class AnalysisExcel<T> implements DataHandleAction<T> {

	private AnalysisExcelForSax<T> analysisExcelForSax = new AnalysisExcelForSax<T>();
	
	private ReflexObject<T> reflexObject = new ReflexObject<T>();
	
	@Override
	public void prepare(ActionContext<T> act) {
		
		if(act.getUseSax()) {
			commit(act);
			return;
		}
		
		InputStream inputStream = act.getInputStream();
		
		Workbook wb = null;
		
		List<SheetInfo<T>> excelInfo = act.getSheets();
		
		boolean distinct = act.getDistinct();
		
		try{
			switch(act.getFileType()){
			case XLS:
				wb = new HSSFWorkbook(inputStream);
				break;
			case XLSX:
				wb = new XSSFWorkbook(inputStream);
				break;
			}
		}catch(IOException e){
			e.printStackTrace();
			rollback(act);
		}finally {
			try {
				if(inputStream != null) {
					inputStream.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		
		int numberOfSheets = 0;
		if(wb != null) {
			numberOfSheets = wb.getNumberOfSheets();
		}
		act.setSheetSize(numberOfSheets);
		int firstSheetHeadNum = -1;
		
		if(numberOfSheets >= 1){
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(0);
			int cellNumber = row.getLastCellNum();
			firstSheetHeadNum = cellNumber;
			
			for(int n=0; n<0; n++){
				int cellNum = 0;
				if(row != null) {
					cellNum = row.getLastCellNum();
				}
				
				for(int i=0; i<cellNum; i++){
					Cell cell = row.getCell(i);
					CellType cellType = CellType._NONE;
					if(cell != null)
						cellType = cell.getCellTypeEnum();
					if(cellType == CellType._NONE || cellType == CellType.BLANK){
						cellNum = i;
					}
				}
				
				if(firstSheetHeadNum == -1){
					firstSheetHeadNum = cellNum;
				}
				
			}
		}
		
		for(int i=0; i<numberOfSheets; i++){
			Sheet sheet = wb.getSheetAt(i);
			SheetInfo<T> sheetInfo = new SheetInfo<T>();
			List<T> analysisSheet = analysisSheet(sheet,act.getObjects(),sheetInfo,distinct);
			if(analysisSheet == null) {
				sheetInfo.setSheetName(sheet.getSheetName());
				excelInfo.add(sheetInfo);
				continue;
			}
			sheetInfo.setInfo(analysisSheet);
			sheetInfo.setSheetName(sheet.getSheetName());
			sheetInfo.setLine(sheet.getLastRowNum());
			sheetInfo.setColumn(firstSheetHeadNum);
			sheetInfo.setCorrectLine(analysisSheet.size());
			sheetInfo.setBlankLineSize(sheetInfo.getBlankLine().size());
			sheetInfo.setErrorLineSize(sheetInfo.getErrorLine().size());
			sheetInfo.setRepeatLineSize(sheetInfo.getRepeatLine().size());
			excelInfo.add(sheetInfo);
		}
		
		commit(act);
		
	}

	@Override
	public boolean commit(ActionContext<T> act) {
		
		analysisExcelForSax.prepare(act);
		
		return true;
	}

	@Override
	public boolean rollback(ActionContext<T> act) {

		act.setResult(false);
		
		return false;
	}
	
	/**
	 * 
	 * Parse sheet into row list
	 * @param sheet
	 * @return   
	 * @return List<String>
	 */
	@SuppressWarnings("deprecation")
	List<T> analysisSheet(Sheet sheet,Class<? extends T> object,SheetInfo<T> sheetInfo, boolean distinct){
		int lastRowNum = sheet.getLastRowNum();
		List<T> sheetList = new ArrayList<T>();
		Row row = sheet.getRow(0);
		if(row == null) {
			return null;
		}
		int cellNum = row.getLastCellNum();
		List<String> headStr = new ArrayList<String>();
		Set<T> distinctSet = new HashSet<T>();
		for(int i=0; i<cellNum; i++){
			Cell cell = row.getCell(i);
			int cellType = -1;
			if(cell != null) {
				cellType = cell.getCellType();
			}
			switch(cellType){
			case Cell.CELL_TYPE_BLANK:
				headStr.add("");
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				headStr.add(cell.getBooleanCellValue() + "");
				break;
			case Cell.CELL_TYPE_FORMULA:
				headStr.add(cell.getCellFormula());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				headStr.add(cell.getNumericCellValue() + "");
				break;
			case Cell.CELL_TYPE_STRING:
				headStr.add(cell.getStringCellValue() + "");
				break;
			case Cell.CELL_TYPE_ERROR:
				break;
			default:
				headStr.add("");
			}
		}
		for(int i=1; i<=lastRowNum; i++){
			
			T t = null;
			
			Row row2 = sheet.getRow(i);
			
			if(row2 == null) {
				sheetInfo.addBlankLine(i);
				continue;
			}
			
			boolean rowEmpty = isRowEmpty(row2);
			
			if(rowEmpty) {
				sheetInfo.addBlankLine(i + 1);
				continue;
			}
			
			List<String> analysisRow = analysisRow(row2,cellNum);
			
			t = reflexObject.getReflexObject(object,headStr,analysisRow,sheet.getSheetName(), i + 1);
			
			if(t == null) {
				sheetInfo.addErrorLine(i + 1);
				continue;
			}
			
			boolean flag = true;
			if(distinct) {
				boolean add = distinctSet.add(t);
				if(!add) {
					sheetInfo.addRepeatLine(i + 1);
					flag = false;
				}
			}
			
			if(flag) {
				sheetList.add(t);
			}
			
		}
		
		return sheetList;
	}
	

	/**
	 * Convert each cell into a string and save it
	 * @param row
	 * @param cellNum
	 * @return
	 */
	@SuppressWarnings("deprecation")
	List<String> analysisRow(Row row,int cellNum){
		List<String> cellLi = new ArrayList<String>();
		for(int y = 0; y < cellNum; y++){
			Cell cell = row.getCell(y);
			int cellType = 3;
			if(cell != null)
				cellType = cell.getCellType();
			switch(cellType){
			case Cell.CELL_TYPE_BLANK:
				cellLi.add("");
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				cellLi.add(cell.getBooleanCellValue() + "");
				break;
			case Cell.CELL_TYPE_FORMULA:
				cellLi.add(String.valueOf(cell.getNumericCellValue()));
				// cellLi.add(cell.getCellFormula());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				short dataFormat = cell.getCellStyle().getDataFormat();
				if(DateUtil.isCellDateFormatted(cell)){
					Date date = cell.getDateCellValue();
					if(dataFormat == 179){
						Date javaDate = DateUtil.getJavaDate(cell.getNumericCellValue());
						SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM");
						String string = simpleDateFormat.format(javaDate);
						cellLi.add(string);
					}else if(dataFormat == 58 || dataFormat == 177) {
						Date javaDate = DateUtil.getJavaDate(cell.getNumericCellValue());
						SimpleDateFormat simpleDateFormat = new SimpleDateFormat("MM-dd");
						String string = simpleDateFormat.format(javaDate);
						cellLi.add(string);
					}else if(date != null){
						SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
						String string = simpleDateFormat.format(date);
						cellLi.add(string);
					}else{
						cellLi.add("1970-01-01");
					}
					break;
				}
				
				DecimalFormat df = new DecimalFormat("#.######");
				String dateString = df.format(cell.getNumericCellValue());
				cellLi.add(dateString);
				
				break;
			case Cell.CELL_TYPE_STRING:
				cellLi.add(cell.getStringCellValue() + "");
				break;
			case Cell.CELL_TYPE_ERROR:
				break;
			default:
				cellLi.add("");
			}
		}	
		return cellLi;
	}
	
	/**
	 * 
	 * Judge whether the line is empty
	 * @param row
	 * @return   
	 * @return boolean
	 */
	@SuppressWarnings("deprecation")
	public boolean isRowEmpty(Row row){
		for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
			
			Cell cell = row.getCell(c);
		
			if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
				return false;
			}
		
		}
		
		return true;
	}
	
}