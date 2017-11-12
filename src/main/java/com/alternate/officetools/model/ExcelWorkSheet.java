package com.alternate.officetools.model;

import org.apache.commons.lang3.ArrayUtils;

import java.util.*;

public class ExcelWorkSheet {
    private String name;
    private Object[] columnNames;
    private Map<String, Object[]> data;
    private Map<String, Object[]> duplicateEntries;
    private Map<String, Integer> groups;
    private int[] keyColumns;
    private int groupingColumn;

    public ExcelWorkSheet(String name, int[] keyColumns, int groupingColumn) {
        this();
        this.setName(name);
        this.keyColumns = keyColumns;
        this.groupingColumn = groupingColumn;
    }

    public ExcelWorkSheet() {
        this.data = new LinkedHashMap<>();
        this.duplicateEntries = new LinkedHashMap<>();
        this.groups = new TreeMap<>();
        this.keyColumns = new int[]{0};
        this.groupingColumn = 0;
    }

    public void insertRow(Object[] row, boolean seperateDuplicates) {
        String key = generateKey(row, this.keyColumns);

        if (!seperateDuplicates) {
            key = "entry_" + (this.data.size() + 1);
            this.data.put(key, row);
        } else if (this.isInDuplicateEntries(key)) {
            Object[] rows = this.duplicateEntries.get(key);
            rows = ArrayUtils.add(rows, row);
            this.duplicateEntries.put(key, rows);
        } else if (this.isInData(key)) {
            Object[] previousRow = this.data.get(key);
            this.data.remove(key);
            this.duplicateEntries.put(key, new Object[]{previousRow, row});
        } else {
            this.data.put(key, row);
            this.insertIntoGroup(row);
        }
    }

    public boolean isInData(String key){
        return this.data.containsKey(key);
    }

    public boolean isInDuplicateEntries(String key){
        return this.duplicateEntries.containsKey(key);
    }

    private void insertIntoGroup(Object[] row) {
        String key = row[this.groupingColumn].toString();

        if (this.groups.containsKey(key)) {
            this.groups.put(key, this.groups.get(key) + 1);
        } else {
            this.groups.put(key, 1);
        }
    }

    private static String generateKey(Object[] row, int[] keyColumns) {
        StringBuilder key = new StringBuilder();

        for (int keyColumn : keyColumns) {
            key.append(row[keyColumn]);
            key.append("_");
        }

        return key.toString();
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

    public Map<String, Object[]> getDuplicateEntries() {
        return duplicateEntries;
    }

    public int[] getKeyColumns() {
        return keyColumns;
    }

    public int getGroupingColumn() {
        return groupingColumn;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setColumnNames(Object[] columnNames) {
        this.columnNames = columnNames;
    }

    public void setKeyColumns(int[] keyColumns) {
        this.keyColumns = keyColumns;
    }

    public void setGroupingColumn(int groupingColumn) {
        this.groupingColumn = groupingColumn;
    }

    //</editor-fold>
}
