package cc.ebatis.util;

import java.io.IOException;
import java.io.InputStream;

import cc.ebatis.emnu.FileType;
import cc.ebatis.pojo.ActionContext;
@SuppressWarnings("rawtypes")
public class CheckFileType {
	
	public static String getByteStr(byte[] bytes){
		StringBuilder hex = new StringBuilder();
		for (int i = 0; i < bytes.length; i++) {  
            String temp = Integer.toHexString(bytes[i] & 0xFF);  
            if (temp.length() == 1) {  
                hex.append("0");  
            }
            hex.append(temp.toLowerCase());  
        }  
        return hex.toString();  
	}
	
	/** 
     * read file header 
     */  
    private static String getFileHeader(InputStream inputStream) throws IOException {  
        byte[] b = new byte[28];  
        inputStream.read(b, 0, 28);  
        
        return getByteStr(b);  
    }  
      
    /** 
     * Judge file type
     */  
    public static FileType getType( ActionContext act) throws IOException {  
        
    	InputStream inputStream = act.getInputStream();
    	
        String fileHead = getFileHeader(inputStream);  
        
        inputStream.close(); 
        inputStream = null;
        
        if (fileHead == null || fileHead.length() == 0) {  
            return null;  
        }  
        fileHead = fileHead.toUpperCase();  
        FileType[] fileTypes = FileType.values();  
        for (FileType type : fileTypes) {  
            if (fileHead.startsWith(type.getValue())) {  
                return type;  
            }  
        }  
        return null;  
    }  
      
	
}
