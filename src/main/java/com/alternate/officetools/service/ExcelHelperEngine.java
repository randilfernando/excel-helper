package com.alternate.officetools.service;

import com.alternate.officetools.model.ExcelWorkBook;
import com.alternate.officetools.model.ExcelWorkSheet;

public class ExcelHelperEngine {
    private ExcelWorkBookReader excelWorkBookReader;
    private ExcelWorkBookWriter excelWorkBookWriter;
    private ExcelWorkSheetMerger excelWorkSheetMerger;

    public ExcelHelperEngine(){
        this.excelWorkBookReader = new ExcelWorkBookReader();
        this.excelWorkBookWriter = new ExcelWorkBookWriter();
        this.excelWorkSheetMerger = new ExcelWorkSheetMerger();
    }

    public ExcelWorkBook readExcelWorkBook(String filePath, boolean useBasePath, int[] keyColumns, int groupingColumn){
        return this.excelWorkBookReader.readExcelWorkBook(filePath, useBasePath, keyColumns, groupingColumn);
    }

    public void writeExcelWorkBook(String filePath, boolean useBasePath, ExcelWorkBook excelWorkBook){
        this.excelWorkBookWriter.writeExcelWorkBook(filePath, useBasePath, excelWorkBook);
    }

    public ExcelWorkBook joinWorkBooksUsingKey(ExcelWorkBook workBook1, ExcelWorkBook workBook2){
        return this.excelWorkSheetMerger.joinWorkBooksUsingKey(workBook1, workBook2);
    }
}
