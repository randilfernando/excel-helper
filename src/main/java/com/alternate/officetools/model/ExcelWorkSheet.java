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
    private int[] groupingColumns;

    public ExcelWorkSheet(String name, int[] keyColumns, int[] groupingColumns) {
        this();
        this.setName(name);
        this.keyColumns = keyColumns;
        this.groupingColumns = groupingColumns;
    }

    public ExcelWorkSheet() {
        this.data = new LinkedHashMap<>();
        this.duplicateEntries = new LinkedHashMap<>();
        this.groups = new TreeMap<>();
        this.keyColumns = new int[]{0};
        this.groupingColumns = new int[]{0};
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
            this.insertIntoGroup(row);
        } else if (this.isInData(key)) {
            Object[] previousRow = this.data.get(key);
            this.data.remove(key);
            this.duplicateEntries.put(key, new Object[]{previousRow, row});
            this.insertIntoGroup(row);
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
        StringBuilder key = new StringBuilder();

        for (int groupingColumn: this.groupingColumns){
            key.append(row[groupingColumn].toString()).append("_");
        }

        if (this.groups.containsKey(key.toString())) {
            this.groups.put(key.toString(), this.groups.get(key.toString()) + 1);
        } else {
            this.groups.put(key.toString(), 1);
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

    public Map<String, Integer> getGroups() {
        return groups;
    }

    public int[] getKeyColumns() {
        return keyColumns;
    }

    public int[] getGroupingColumn() {
        return groupingColumns;
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

    public void setGroupingColumn(int[] groupingColumns) {
        this.groupingColumns = groupingColumns;
    }

    //</editor-fold>
}
