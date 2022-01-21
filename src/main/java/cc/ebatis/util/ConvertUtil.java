package cc.ebatis.util;

/**
 * 转换工具
 * @author Administrator
 *
 */
public class ConvertUtil {
	
	/**
	 * Capitalize string initials
	 * @param str
	 * @return
	 */
	public static String upperCase(String str) {  
	    char[] ch = str.toCharArray();  
	    if (ch[0] >= 'a' && ch[0] <= 'z') {  
	        ch[0] = (char) (ch[0] - 32);  
	    }  
	    return new String(ch);  
	}
}
