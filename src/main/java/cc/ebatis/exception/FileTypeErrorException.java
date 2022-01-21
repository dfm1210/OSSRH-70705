package cc.ebatis.exception;

/**
 * File type exception
 * @author Steve
 *
 */
public class FileTypeErrorException extends Exception {
	
	
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public FileTypeErrorException(){
		super();
	}
	
	public FileTypeErrorException(String message){
		super(message);
	}
	
	
}
