package com.alternate.officetools.model;

import java.util.Map;
import java.util.TreeMap;

public class ExcelWorkSheet {
    private String name;
    private Object[] columnNames;
    private Map<String, Object[]> data;

    public ExcelWorkSheet(String name){
        this();
        this.setName(name);
    }

    public ExcelWorkSheet(){
        this.data = new TreeMap<>();
    }

    public void insertRow(String key, Object[] row){
        this.data.put(key, row);
    }

    //<editor-fold> desc="Getters and Setters"

    public String getName() {
        return name;
    }

    public Object[] getColumnNames() {
        return columnNames;
    }

    public Map<String, Object[]> getData() {
        return data;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setColumnNames(Object[] columnNames) {
        this.columnNames = columnNames;
    }

    //</editor-fold>
}
