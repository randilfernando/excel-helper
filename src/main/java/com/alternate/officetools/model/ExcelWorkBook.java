package com.alternate.officetools.model;

import com.alternate.officetools.helper.ExcelHelper;

import java.util.ArrayList;
import java.util.Date;

public class ExcelWorkBook {
    private String name;
    private final ArrayList<ExcelWorkSheet> excelWorkSheets;

    public ExcelWorkBook(String name) {
        this();
        this.name = name;
    }

    public ExcelWorkBook() {
        this.name = ExcelHelper.generateNameUsingDate("Book");
        this.excelWorkSheets = new ArrayList<ExcelWorkSheet>();
    }

    public void addWorkSheet(ExcelWorkSheet excelWorkSheet){
        if (excelWorkSheet.getName() == null) {
            String name = ExcelHelper.generateNameUsingWorkSheetNumber("Sheet", this.excelWorkSheets.size() + 1);
            excelWorkSheet.setName(name);
        }
        this.excelWorkSheets.add(excelWorkSheet);
    }

    public void removeWorkSheet(ExcelWorkSheet excelWorkSheet){
        this.excelWorkSheets.remove(excelWorkSheet);
    }

    public void removeWorkSheet(String name){
        int id = -1;
        for (int i = 0; i < this.excelWorkSheets.size(); i++){
            ExcelWorkSheet excelWorkSheet = this.excelWorkSheets.get(i);
            if (excelWorkSheet.getName().equals(name)){
                id = i;
                break;
            }
        }

        if (id != -1) {
            this.excelWorkSheets.remove(id);
        }
    }

    public void clearWorkBook(){
        this.excelWorkSheets.clear();
    }

    //<editor-fold> desc="Getters and Setters"

    public String getName() {
        return name;
    }

    public ArrayList<ExcelWorkSheet> getExcelWorkSheets() {
        return excelWorkSheets;
    }

    public void setName(String name) {
        this.name = name;
    }

    //</editor-fold
}
