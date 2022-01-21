package cc.ebatis.pojo;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import cc.ebatis.emnu.FileType;

public class ActionContext<T> {
	
	private File file;
	
	private ByteArrayOutputStream byteArrayOutputStream = null;
	
	private List<SheetInfo<T>> sheets = new ArrayList<SheetInfo<T>>();
	
	private FileType fileType;
	
	private Integer SheetSize;
	
	private Integer fileSizeByte;
	
	private Class<? extends T> objects;
	
	private boolean useSax = false;
	
	private boolean distinct = false;
	
	private boolean result = false;

	public boolean getDistinct() {
		return distinct;
	}

	public void setDistinct(boolean distinct) {
		this.distinct = distinct;
	}

	public boolean getUseSax() {
		return useSax;
	}

	public void setUseSax(boolean useSax) {
		this.useSax = useSax;
	}

	public boolean getResult() {
		return result;
	}

	public void setResult(boolean result) {
		this.result = result;
	}

	public Class<? extends T> getObjects() {
		return objects;
	}

	public void setObjects(Class<? extends T> objects) {
		this.objects = objects;
	}

	public Integer getFileSizeByte() {
		return fileSizeByte;
	}

	public void setFileSizeByte(Integer fileSizeByte) {
		this.fileSizeByte = fileSizeByte;
	}

	public Integer getSheetSize() {
		return SheetSize;
	}

	public void setSheetSize(Integer sheetSize) {
		SheetSize = sheetSize;
	}

	public FileType getFileType() {
		return fileType;
	}

	public void setFileType(FileType fileType) {
		this.fileType = fileType;
	}
	
	public File getFile() {
		return file;
	}

	public void setFile(File file) {
		this.file = file;
	}

	public List<SheetInfo<T>> getSheets() {
		return sheets;
	}

	public void setSheets(List<SheetInfo<T>> sheets) {
		this.sheets = sheets;
	}
	
	public void addSheets(SheetInfo<T> sheets) {
		if(this.sheets != null) {
			this.sheets.add(sheets);
		}
		
	}

	public InputStream getInputStream() {
		if(this.file != null) {
			InputStream inputStream = null;
			try {
				inputStream = new FileInputStream(this.file);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			}
			return inputStream;
		}
		
		InputStream inputStreamReal = new ByteArrayInputStream(this.byteArrayOutputStream.toByteArray());
		
		return inputStreamReal;
	}
	
	public void setInputStream(InputStream inputStream) throws IOException {
		
		if(this.byteArrayOutputStream == null) {
			this.byteArrayOutputStream = new ByteArrayOutputStream();
			byte[] buffer = new byte[inputStream.available()];
			
			int len;  
			while ((len = inputStream.read(buffer)) > -1 ) {  
			    this.byteArrayOutputStream.write(buffer, 0, len);  
			}
			this.byteArrayOutputStream.flush();
		}
	}

	@Override
	public String toString() {
		return "{\"sheets\":" + sheets + ", \"fileType\":\"" + fileType + "\", \"SheetSize\":" + SheetSize
				+ ", \"fileSizeByte\":" + fileSizeByte + ", \"useSax\":" + useSax + ", \"distinct\":"
				+ distinct + ", \"result\":" + result + "} ";
	}


}