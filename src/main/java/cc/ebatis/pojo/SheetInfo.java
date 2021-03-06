package cc.ebatis.pojo;

import java.util.ArrayList;
import java.util.List;

public class SheetInfo<T> {
	
	private Integer line;
	
	private String sheetName;
	
	private Integer column;
	
	private Integer correctLine;
	
	private Integer blankLineSize;
	
	private Integer errorLineSize;
	
	private Integer repeatLineSize;	
	
	private List<T> info;
	
	private List<Integer> blankLine;
	
	private List<Integer> errorLine;
	
	private List<Integer> repeatLine;
	
	public SheetInfo(){
		this.line = 0;
		this.column = 0;
		this.correctLine = 0;
		this.blankLineSize = 0;
		this.errorLineSize = 0;
		this.repeatLineSize = 0;
		this.info = new ArrayList<T>();
		this.blankLine = new ArrayList<Integer>();
		this.errorLine = new ArrayList<Integer>();
		this.repeatLine = new ArrayList<Integer>();
	}
	
	
	public Integer getBlankLineSize() {
		return blankLineSize;
	}


	public void setBlankLineSize(Integer blankLineSize) {
		this.blankLineSize = blankLineSize;
	}


	public Integer getErrorLineSize() {
		return errorLineSize;
	}


	public void setErrorLineSize(Integer errorLineSize) {
		this.errorLineSize = errorLineSize;
	}


	public Integer getRepeatLineSize() {
		return repeatLineSize;
	}


	public void setRepeatLineSize(Integer repeatLineSize) {
		this.repeatLineSize = repeatLineSize;
	}


	public List<Integer> getRepeatLine() {
		return repeatLine;
	}


	public void addRepeatLine(Integer repeatLine) {
		this.repeatLine.add(repeatLine);
	}


	public List<Integer> getErrorLine() {
		return errorLine;
	}

	public void addErrorLine(Integer errorLine) {
		this.errorLine.add(errorLine);
	}

	public List<Integer> getBlankLine() {
		return blankLine;
	}

	public void addBlankLine(Integer blankLine) {
		this.blankLine.add(blankLine);
	}

	public Integer getCorrectLine() {
		return correctLine;
	}

	public void setCorrectLine(Integer correctLine) {
		this.correctLine = correctLine;
	}

	public Integer getLine() {
		return line;
	}

	public void setLine(Integer line) {
		this.line = line;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public Integer getColumn() {
		return column;
	}

	public void setColumn(Integer column) {
		this.column = column;
	}

	public List<T> getInfo() {
		return info;
	}

	public void setInfo(List<T> info) {
		this.info = info;
	}


	@Override
	public String toString() {
		return "{\"line\":" + line + ", \"sheetName\":\"" + sheetName + "\", \"column\":" + column
				+ ", \"correctLine\":" + correctLine + ", \"blankLineSize\":" + blankLineSize
				+ ", \"errorLineSize\":" + errorLineSize + ", \"repeatLineSize\":" + repeatLineSize
				+ ", \"info\":" + info + ", \"blankLine\":" + blankLine + ", \"errorLine\":" + errorLine
				+ ", \"repeatLine\":" + repeatLine + "} ";
	}

	
}
