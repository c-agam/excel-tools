package com.mackey.excel.eventmodel;

import java.io.File;
import java.io.InputStream;

public interface IReader {
	void process(String path) throws Exception;

	void process(File file) throws Exception;

	void process(InputStream in) throws Exception;
}
