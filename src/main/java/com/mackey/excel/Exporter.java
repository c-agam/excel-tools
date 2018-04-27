package com.mackey.excel;

import com.alibaba.fastjson.JSONArray;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.io.IOException;

/**
 * @author zul
 */
public interface Exporter {
    void createSheet();
    void createHeader(SXSSFSheet sheet);
    void exporter(JSONArray data) throws Exception;
    void write() throws IOException;
}
