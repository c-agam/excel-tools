package com.mackey.excel;

import java.math.BigDecimal;
import java.util.Date;

public interface DataFormat {
	/**
	 * 时间格式化yyyy-MM-dd
	 * @param time
	 * @param format
	 * @return
	 */
	public String dateFormat(Date value);
	
	/**
	 * 时间格式化yyyy-MM-dd HH:mm:ss
	 * @param time
	 * @param format
	 * @return
	 */
	public String dateFormats(Date value);
	
	/**
	 * long格式化 yyyy-MM-dd HH:mm:ss
	 * @param value
	 * @return
	 */
	public String longFormat(Long value);

	/**
	 * 时间格式化yyyy-MM-dd
	 * @param time
	 * @param format
	 * @return
	 */
	public String longFormats(Long value);
	/**
	 * Double格式化
	 * @param value
	 * @return
	 */
	public String doubleFormat(Double value);
	
	/**
	 * BigDecimal格式化
	 * @param bigDecimal
	 * @return
	 */
	public String bigDecimalFormat(BigDecimal value);
	
	/**
	 * Integer格式化
	 * @param value
	 * @return
	 */
	public String integerFormat(Integer value);
	
	/**
	 * long格式化 yyyy-MM-dd HH:mm:ss
	 * @param value
	 * @return
	 */
	public String stringFormat(String value);

	/**
	 * 时间格式化yyyy-MM-dd
	 * @param time
	 * @param format
	 * @return
	 */
	public String stringFormats(String value);
	
	/**
	 * 清理内存缓存数据
	 */
	public void clear();
}
