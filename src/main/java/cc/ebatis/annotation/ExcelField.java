package cc.ebatis.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Specify the location of the attributes and attribute names contained 
 * in the generated excel file, whether they have been merged, and the order
 * @author Steve
 *
 */
@Documented
@Target(value = {ElementType.FIELD})
@Retention(value = RetentionPolicy.RUNTIME)
public @interface ExcelField {
	
	public int position();
	
	public String name() default "";
	
	public int width() default -1;
	
	public boolean merge() default false;
	
}
