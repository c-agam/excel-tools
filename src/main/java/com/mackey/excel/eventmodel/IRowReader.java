package com.mackey.excel.eventmodel;

import java.util.List;

public interface IRowReader {
	void row(int sheetIndex, int rowIndex, List<String> rowData);
}
