package cc.ebatis.impl;

import java.io.IOException;
import java.io.InputStream;

import cc.ebatis.api.DataHandleAction;
import cc.ebatis.emnu.FileType;
import cc.ebatis.exception.FileTypeErrorException;
import cc.ebatis.pojo.ActionContext;
import cc.ebatis.util.CheckFileType;

/**
 * 
 * Verify table type and size
 * @author Steve
 *
 */
public class VerificationTable<T> implements DataHandleAction<T>{

	AnalysisExcel<T> analysisExcel = new AnalysisExcel<T>();
	
	@Override
	public void prepare(ActionContext<T> act) {
		
		FileType type = null;
		InputStream inputStream = null;
		try {
			type = CheckFileType.getType(act);
			if(type == null){
				throw new FileTypeErrorException("This file type is error, Not's xsl or xslx");
			}
			
			act.setFileType(type);
			
			inputStream = act.getInputStream();
			int available = inputStream.available();
			
			if(type == FileType.XLSX){
				act.setUseSax(true);
			}
			act.setFileSizeByte(available);
			
			commit(act);
			
		} catch (IOException e) {
			e.printStackTrace();
		} catch (FileTypeErrorException e) {
			e.printStackTrace();
		} finally{
			try {
				if(inputStream != null) {
					inputStream.close();
				}
				
			} catch (IOException e) {
				e.printStackTrace();
			}
			inputStream = null;
		}
		
	}

	@Override
	public boolean commit(ActionContext<T> act) {

		analysisExcel.prepare(act);
		
		return true;
	}

	@Override
	public boolean rollback(ActionContext<T> act) {
		
		act.setResult(false);
		
		return false;
	}
	
}