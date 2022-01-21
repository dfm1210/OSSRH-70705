package cc.ebatis.exception;

/**
 * no head exception
 * @author Steve
 *
 */
public class NoHeaderException extends Exception {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public NoHeaderException(){
		super();
	}
	
	public NoHeaderException(String message){
		super(message);
	}
}
