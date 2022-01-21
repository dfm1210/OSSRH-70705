package cc.ebatis.exception;

/**
 * Enableexcelmaker annotation not found in entity
 * @author Steve
 *
 */
public class NoEnableExcelMakerException extends Exception {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public NoEnableExcelMakerException(){
		super();
	}
	
	public NoEnableExcelMakerException(String message){
		super(message);
	}
	
	
}
