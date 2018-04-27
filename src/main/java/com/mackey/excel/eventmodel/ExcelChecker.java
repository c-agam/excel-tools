package com.mackey.excel.eventmodel;

import java.util.List;

/**
 * 
 * excel校验
 * @author zul
 *
 */
public interface ExcelChecker {
	/**
	 * 根据模板校验表头
	 * @param headerTemplate
	 * @param header
	 * @return
	 */
	public boolean checkHeader(List<String> headerTemplate, List<String> header);
	
	/**
	 * 校验行数据
	 * @param rowData
	 * @param header
	 * @return
	 */
	public boolean checkRow(List<String> rowData, List<String> header);
}
