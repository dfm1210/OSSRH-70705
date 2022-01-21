package cc.ebatis.util;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import cc.ebatis.annotation.LineNumber;
import cc.ebatis.annotation.Mapping;
import cc.ebatis.annotation.MappingSheetName;

/**
 * Reflection object tool
 * @author Administrator
 *
 * @param <T>
 */
public class ReflexObject<T> {
	
	private Field[] fields = null;
	private Class<Mapping> mapping = Mapping.class;
	private Class<MappingSheetName> mappingSheetName = MappingSheetName.class;
	private Class<LineNumber> lineNumber = LineNumber.class;
	
	/**
	 * Reflect the cell information into the java bean
	 * @param class1 
	 * @param headStr
	 * @param analysisRow
	 * @param sheetName
	 * @param lineNum
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public T getReflexObject(Class<? extends T> class1, List<String> headStr, List<String> analysisRow, String sheetName, int lineNum){
		
		Object object = null;
		
		try {
			
			Constructor<? extends T> constructor = class1.getConstructor();
			object = constructor.newInstance();
			
			if(fields == null) {
				fields = class1.getDeclaredFields();
			}
			
			for(Field x : fields) {
				Mapping m = x.getAnnotation(mapping);
				MappingSheetName msn = x.getAnnotation(mappingSheetName);
				LineNumber ln = x.getAnnotation(lineNumber);
				
				if(m != null) {
					boolean mappingOperation = this.mappingOperation(class1, object, x, m, headStr, analysisRow);
					if(m.delNull() && !mappingOperation) {
						return null;
					}
					
				}
				
				if(msn != null) {
					this.sheetNameOperation(class1, object, x, sheetName);		
				}
				
				if(ln != null) {
					this.lineNumberOperation(class1, object, x, lineNum);
				}

			}
		
		}catch (IllegalAccessException e) {
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			e.printStackTrace();
		} catch (InvocationTargetException e) {
			e.printStackTrace();
		} catch (NoSuchMethodException e) {
			e.printStackTrace();
		} catch (SecurityException e) {
			e.printStackTrace();
		} catch (InstantiationException e) {
			e.printStackTrace();
		}
		
		return (T)object;
		
	}
	
	
	/**
	 * Field mapping operation
	 * @param object
	 * @param headStr
	 * @param analysisRow
	 * @throws SecurityException 
	 * @throws NoSuchMethodException 
	 * @throws InvocationTargetException 
	 * @throws IllegalArgumentException 
	 * @throws IllegalAccessException 
	 */
	private boolean mappingOperation(Class<?> class1, Object object, Field field, Mapping mapping, List<String> headStr, List<String> analysisRow) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException {
		
		String fieldName = field.getName();
		String title = mapping.key();
		String rex = mapping.rex();
		int length = mapping.length();
		String methodName = new StringBuilder()
				.append("set").append(ConvertUtil.upperCase(fieldName)).toString();
		
		String fieldType = field.getType().toString();
		
		for(int y=0; y<headStr.size(); y++) {
			String thisHead = headStr.get(y);
			if(thisHead != null && title.equals(thisHead) && !thisHead.equals("")){
				
				String string = analysisRow.get(y);
				if(length > 0 && string.length() > length){
					string = string.substring(0, length);
				}
				
				if(!rex.equals("")) {
					
					Pattern compile = Pattern.compile(rex);
					Matcher matcher = compile.matcher(string);
					if(!matcher.matches()) {
						return false;
					}
				}
				
				this.screening(class1, object, methodName, fieldType, string);
				
				return true;
			}
		}
		
		return false;
	}
	
	/**
	 * Sheet name mapping operation
	 * @param class1
	 * @param object
	 * @param x
	 * @param m
	 * @param sheetName
	 * @throws SecurityException 
	 * @throws NoSuchMethodException 
	 * @throws InvocationTargetException 
	 * @throws IllegalArgumentException 
	 * @throws IllegalAccessException 
	 */
	private void sheetNameOperation(Class<?> class1, Object object, Field field, String sheetName) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException{
		String fieldName = field.getName();

		String methodName = new StringBuilder()
				.append("set").append(ConvertUtil.upperCase(fieldName)).toString();
		
		String fieldType = field.getType().toString();
		
		this.screening(class1, object, methodName, fieldType, sheetName);
	}
	
	/**
	 * Mapping rows to entities
	 * @param class1
	 * @param object
	 * @param field
	 * @param lineNum
	 * @throws SecurityException 
	 * @throws NoSuchMethodException 
	 * @throws InvocationTargetException 
	 * @throws IllegalArgumentException 
	 * @throws IllegalAccessException 
	 */
	private void lineNumberOperation(Class<?> class1, Object object, Field field, int lineNum) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException{
		String fieldName = field.getName();

		String methodName = new StringBuilder()
				.append("set").append(ConvertUtil.upperCase(fieldName)).toString();
		
		String fieldType = field.getType().toString();
		
		this.screening(class1, object, methodName, fieldType, String.valueOf(lineNum));
	}
	
	/**
	 * Filter assignments by type
	 * @param object
	 * @Param methodName
	 * @param fieldType
	 * @param fieldValue
	 * @throws SecurityException 
	 * @throws NoSuchMethodException 
	 * @throws InvocationTargetException 
	 * @throws IllegalArgumentException 
	 * @throws IllegalAccessException 
	 */
	private void screening(Class<?> class1, Object object, String methodName, String fieldType, String fieldValue) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException {
		
		Double parseToDouble = null;
		
		switch(fieldType){
		case "class java.util.Date":
			SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
			format.setLenient(false);
			Date parse = null;
			try{
				parse = format.parse(fieldValue);
			}catch(ParseException e){
				parse = null;
			}
			class1.getMethod(methodName, Date.class).invoke(object, parse);
			break;
		case "class java.lang.Integer":
			Integer parseInt = null;
			
			parseToDouble = parseToDouble(fieldValue);
			if(parseToDouble != null) {
				double doubleVal = parseToDouble;
				parseInt = (int)doubleVal;
			}
			try {
				parseInt = Integer.parseInt(fieldValue);
			}catch(NumberFormatException e) {}
			
			class1.getMethod(methodName, Integer.class).invoke(object, parseInt);
			break;
		case "class java.lang.String":
			if(fieldValue.equals("")) {
				break;
			}
			class1.getMethod(methodName, String.class).invoke(object, fieldValue);
			break;
		case "class java.lang.Long":
			Long parseLong = null;
			
			parseToDouble = parseToDouble(fieldValue);
			if(parseToDouble != null) {
				double doubleVal = parseToDouble;
				parseLong = (long)doubleVal;
			}
			
			try{
				parseLong = Long.parseLong(fieldValue);
			}catch(NumberFormatException e){}
			class1.getMethod(methodName, Long.class).invoke(object, parseLong);
			break;
		case "class java.lang.Double":
			Double parseDouble = null;
			try{
				parseDouble = Double.parseDouble(fieldValue);
			}catch(NumberFormatException e){}
			class1.getMethod(methodName, Double.class).invoke(object, parseDouble);
			break;
		case "class java.lang.Short":
			Short parseShort = null;
			
			parseToDouble = parseToDouble(fieldValue);
			if(parseToDouble != null) {
				double doubleVal = parseToDouble;
				parseShort = (short)doubleVal;
			}
			
			try{
				parseShort = Short.parseShort(fieldValue);
			}catch(NumberFormatException e){}
			class1.getMethod(methodName, Short.class).invoke(object, parseShort);
			break;
		case "class java.lang.Boolean":
			Boolean parseBoolean = null;
			parseBoolean = Boolean.parseBoolean(fieldValue);
			class1.getMethod(methodName, Boolean.class).invoke(object, parseBoolean);
			break;
		}
	}
	
	/**
	 * @return
	 */
	public Double parseToDouble(String fieldValue) {
		Double parseDouble = null;
		try{
			if(fieldValue != null && fieldValue.contains(".")) {
				parseDouble = Double.parseDouble(fieldValue);
			}
			
		}catch(NumberFormatException e){}
		
		return parseDouble;
	}
}
