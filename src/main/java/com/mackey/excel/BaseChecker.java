package com.mackey.excel;

import com.alibaba.fastjson.JSONArray;


/**
 * @author zul
 */
public interface BaseChecker {
    /**
     * sheet max record
     */
    int max = 1000000;

    void beyondSheetMax(JSONArray data);
}
