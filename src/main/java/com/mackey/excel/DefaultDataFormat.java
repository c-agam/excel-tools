package com.mackey.excel;


import com.mackey.excel.util.IConstants;

import java.text.SimpleDateFormat;
import java.util.Date;

public class DefaultDataFormat extends AbstractDataFormat {
	@Override
	public String longFormat(Long value) {
		String result = "";
		if (value != null) {
			SimpleDateFormat sdf = new SimpleDateFormat(IConstants.FORMAT_DATETIME);
			Date date = new Date(value);
			result = sdf.format(date);
		}
		return result;
	}

	@Override
	public String longFormats(Long value) {
		String result = "";
		if (value != null) {
			SimpleDateFormat sdf = new SimpleDateFormat(IConstants.FORMAT_DATE);
			Date date = new Date(value);
			result = sdf.format(date);
		}
		return result;
	}

	@Override
	public String dateFormat(Date value) {
		String result = "";
		if(value != null) {
			SimpleDateFormat sdf = new SimpleDateFormat(IConstants.FORMAT_DATE);
			result = sdf.format(value);
		}
		return result;
	}
	
	@Override
	public String dateFormats(Date value) {
		String result = "";
		if(value != null) {
			SimpleDateFormat sdf = new SimpleDateFormat(IConstants.FORMAT_DATETIME);
			result = sdf.format(value);
		}
		return result;
	}
	
	@Override
	public String stringFormat(String value) {
		String result = "";
		if (value != null) {
			SimpleDateFormat sdf = new SimpleDateFormat(IConstants.FORMAT_DATETIME);
			Date date = new Date(Long.parseLong(value));
			result = sdf.format(date);
		}
		return result;
	}

	@Override
	public String stringFormats(String value) {
		String result = "";
		if (value != null) {
			SimpleDateFormat sdf = new SimpleDateFormat(IConstants.FORMAT_DATE);
			Date date = new Date(Long.parseLong(value));
			result = sdf.format(date);
		}
		return result;
	}
}
