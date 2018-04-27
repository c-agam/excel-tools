package com.mackey.excel;


import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.mackey.excel.annotation.Excel;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.Closeable;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.sql.Timestamp;
import java.util.Date;

public class CommonExporter implements Exporter , BaseChecker , Closeable {
    /**
     * sheet own record count
     */
    private Integer count = null;

    /**
     * sheet number
     */
    private int sheetCount = 0;

    /**
     * 工作簿
     */
    private SXSSFWorkbook workbook = new SXSSFWorkbook(1000);

    /**
     * sheet表
     */
    private SXSSFSheet sheet;

    /**
     * 数据模型
     */
    private Class<?> model;

    /**
     * 格式化
     */
    private DataFormat dataFormat;

    /**
     * 数据流向
     */
    private FileOutputStream output = null;

    public CommonExporter(FileOutputStream output, Class<?> model) {
        if(output == null || model == null) {
            throw new IllegalArgumentException("output,model are not null...");
        }
        this.output = output;
        this.model = model;
    }

    public CommonExporter(FileOutputStream output, Class<?> model, DataFormat dataFormat) {
        if(output == null || model == null) {
            throw new IllegalArgumentException("output,model are not null...");
        }
        this.output = output;
        this.model = model;
        this.dataFormat = dataFormat;
    }


    private void check(Object[] paramType,Object[] params) {
        if(paramType == null) {
            if(params != null && params.length != 0) {
                throw new IllegalArgumentException("paramType,params must be syn...");
            }
        } else {
            if(params == null) {
                if(paramType.length != 0) {
                    throw new IllegalArgumentException("paramType,params must be syn...");
                }
            } else {
                if(paramType.length != params.length) {
                    throw new IllegalArgumentException("paramType,params must be syn...");
                }
            }
        }
    }

    private void vs(JSONObject jo, Class<?> type, int u, String key,
                    Class<?>[] paramType, Object[] vs) {
        if(type.isAssignableFrom(Date.class)) {
            paramType[u] = Date.class;
            vs[u] = jo.getDate(key);
        }

        if(type.isAssignableFrom(Integer.class)) {
            paramType[u] = Integer.class;
            vs[u] = jo.getInteger(key);
        }

        if(type.isAssignableFrom(String.class)) {
            paramType[u] = String.class;
            vs[u] = jo.getString(key);
        }

        if(type.isAssignableFrom(Double.class)) {
            paramType[u] = Double.class;
            vs[u] = jo.getDouble(key);
        }

        if(type.isAssignableFrom(Long.class)) {
            paramType[u] = Long.class;
            vs[u] = jo.getLong(key);
        }

        if(type.isAssignableFrom(Float.class)) {
            paramType[u] = Float.class;
            vs[u] = jo.getFloat(key);
        }
        if(type.isAssignableFrom(Boolean.class)) {
            paramType[u] = Boolean.class;
            vs[u] = jo.getBoolean(key);
        }
        if(type.isAssignableFrom(Timestamp.class)) {
            paramType[u] = Timestamp.class;
            vs[u] = jo.getTimestamp(key);
        }
    }

    private void getValueByType(Object[] vs, int u, Class<?> type,
                                String value) {
        if(type.isAssignableFrom(Integer.class)) {
            vs[u] = Integer.parseInt(value);
        }
        if(type.isAssignableFrom(String.class)) {
            vs[u] = value;
        }
        if(type.isAssignableFrom(Double.class)) {
            vs[u] = Double.parseDouble(value);
        }
        if(type.isAssignableFrom(Long.class)) {
            vs[u] = Long.parseLong(value);
        }
        if(type.isAssignableFrom(Float.class)) {
            vs[u] = Float.parseFloat(value);
        }
    }

    private void setValue(SXSSFRow row, JSONObject jo, int cellCount,
                          Class<?> type, String field) {
        if(type.isAssignableFrom(Date.class)) {
            try {
                row.createCell(cellCount).setCellValue(jo.getDate(field));;
            } catch (Exception e) {
                row.createCell(cellCount).setCellValue("");
            }
        }

        if(type.isAssignableFrom(Long.class)) {
            try {
                row.createCell(cellCount).setCellValue(jo.getLong(field));;
            } catch (Exception e) {
                row.createCell(cellCount).setCellValue("");
            }
        }

        if(type.isAssignableFrom(Double.class)) {
            try {
                row.createCell(cellCount).setCellValue(jo.getDouble(field));;
            } catch (Exception e) {
                row.createCell(cellCount).setCellValue("");
            }
        }

        if(type.isAssignableFrom(Integer.class)) {
            try {
                row.createCell(cellCount).setCellValue(jo.getInteger(field));;
            } catch (Exception e) {
                row.createCell(cellCount).setCellValue("");
            }
        }

        if(type.isAssignableFrom(Boolean.class)) {
            try {
                row.createCell(cellCount).setCellValue(jo.getBoolean(field));;
            } catch (Exception e) {
                row.createCell(cellCount).setCellValue("");
            }
        }

        if(type.isAssignableFrom(Float.class)) {
            try {
                row.createCell(cellCount).setCellValue(jo.getFloat(field));;
            } catch (Exception e) {
                row.createCell(cellCount).setCellValue("");
            }
        }

        if(type.isAssignableFrom(Timestamp.class)) {
            try {
                row.createCell(cellCount).setCellValue(jo.getTimestamp(field));;
            } catch (Exception e) {
                row.createCell(cellCount).setCellValue("");
            }
        }

        if(type.isAssignableFrom(String.class)) {
            try {
                row.createCell(cellCount).setCellValue(jo.getString(field));;
            } catch (Exception e) {
                row.createCell(cellCount).setCellValue("");
            }
        }
    }

    public void createSheet() {
        sheet = workbook.createSheet("sheet" + sheetCount);
        createHeader(sheet);
        sheetCount++;
    }

    public void createHeader(SXSSFSheet sheet) {
        Field[] fields = model.getDeclaredFields();
        if(fields == null || fields.length == 0) {
            throw new IllegalArgumentException("model is invalid");
        }
        SXSSFRow sxssFRow = sheet.createRow(0);
        for(int u = 0; u < fields.length; u++) {
            Excel excel = fields[u].getAnnotation(Excel.class);
            if(excel != null) {
                sxssFRow.createCell(u).setCellValue(excel.header());
            }
        }
    }

    public void exporter(JSONArray data) throws Exception {
        if(count == null) {
            createSheet();
            count = 0;
        }

        beyondSheetMax(data);

        Field[] fields = model.getDeclaredFields();
        if(fields == null) {
            fields = new Field[]{};
        }

        for(int r = 0; r < data.size(); r++) {
            count ++;

            if(count > max) {
                count = count - max;
                createSheet();
            }

            SXSSFRow row = sheet.createRow(count);
            JSONObject jo = data.getJSONObject(r);

            int cellCount = 0;

            for(int j = 0; j < fields.length; j++) {
                Excel excel = fields[j].getAnnotation(Excel.class);

                if(excel != null) {
                    Class<?> type = fields[j].getType();
                    boolean isFormat = excel.isFormat();
                    String field = excel.field();
                    Class<?> format = excel.format();
                    Class<?>[] defaultParamType = excel.defaultParamType();
                    String[] defaultValue = excel.defaultValue();
                    Class<?>[] paramType = null;
                    String[] params = null;
                    Object[] vs = null;

                    check(defaultParamType,defaultValue);//校验默认参数

                    if(isFormat && !format.isAssignableFrom(NotNeedFormat.class)) {//数据需要格式化

                        if(dataFormat == null) {
                            dataFormat = (DataFormat)format.getConstructor().newInstance();
                        }

                        paramType = excel.paramType();//格式化参数类型
                        params = excel.params();//格式化参数列表

                        check(paramType, params);//校验

                        if(params == null || params.length == 0) {
                            if(defaultValue == null || defaultValue.length == 0) {
                                paramType = new Class<?>[1];
                                vs = new Object[1];
                                vs(jo, type, 0, field, paramType, vs);
                            } else {
                                paramType = new Class<?>[1 + defaultParamType.length];
                                vs = new Object[1 + defaultValue.length];
                                for(int u = 0; u < paramType.length; u++) {
                                    if(u == 0) {
                                        vs(jo, type, 0, field, paramType, vs);
                                    } else {
                                        paramType[u] = defaultParamType[u-1];
                                        getValueByType(vs, u, defaultParamType[u -1], (defaultValue[u -1]));
                                    }
                                }
                            }
                        } else {
                            if(defaultValue == null || defaultValue.length == 0) {
                                vs = new Object[paramType.length];
                                for(int u = 0; u < paramType.length; u++) {
                                    vs(jo,paramType[u],u,params[u],paramType,vs);
                                }
                            } else {
                                Class<?>[] tType = new Class<?>[paramType.length + defaultParamType.length];
                                vs = new Object[paramType.length + defaultParamType.length];

                                for(int u = 0; u < paramType.length; u++) {
                                    tType[u] = paramType[u];
                                    vs(jo,paramType[u],u,params[u],paramType,vs);
                                }

                                for(int u = 0; u < defaultParamType.length; u++) {
                                    int index = u + paramType.length;
                                    tType[index] = defaultParamType[u];
                                    getValueByType(vs, index, defaultParamType[u], defaultValue[u]);
                                }

                                paramType = tType;
                            }
                        }

                        try {
                            row.createCell(cellCount).setCellValue((String)format.getMethod(excel.method(), paramType).invoke(dataFormat, vs));;
                        } catch (Exception e) {
                            row.createCell(cellCount).setCellValue("");
                        }
                    } else {
                        setValue(row, jo, cellCount, type, field);
                    }
                    cellCount++;
                }
            }

        }
    }

    public void beyondSheetMax(JSONArray data) {
        if(data == null) {
            data = new JSONArray();
        }

        if(data.size() > max) {
            throw new IllegalArgumentException("data must less then " + max + " records...");
        }
    }

    public void write() throws IOException {
        if(workbook != null) {
            if(output != null) {
                workbook.write(output);
            }
        }
    }

    public void close() throws IOException {
        if(workbook != null) {
            workbook.close();
        }

        if(output != null) {
            output.close();
        }

        if(dataFormat != null) {
            dataFormat.clear();
        }
    }
}
