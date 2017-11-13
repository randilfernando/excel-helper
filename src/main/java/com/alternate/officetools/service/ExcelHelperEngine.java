package com.alternate.officetools.service;

import com.alternate.officetools.model.ExcelWorkBook;

public class ExcelHelperEngine {
    private ExcelWorkBookReader excelWorkBookReader;
    private ExcelWorkBookWriter excelWorkBookWriter;
    private ExcelWorkSheetMerger excelWorkSheetMerger;
    private ExcelReportGenerator excelReportGenerator;

    public ExcelHelperEngine(){
        this.excelWorkBookReader = new ExcelWorkBookReader();
        this.excelWorkBookWriter = new ExcelWorkBookWriter();
        this.excelWorkSheetMerger = new ExcelWorkSheetMerger();
        this.excelReportGenerator = new ExcelReportGenerator();
    }

    public ExcelWorkBook readExcelWorkBook(String filePath, boolean useBasePath, int[] keyColumns, int[] groupingColumns){
        return this.excelWorkBookReader.readExcelWorkBook(filePath, useBasePath, keyColumns, groupingColumns);
    }

    public void writeExcelWorkBook(String filePath, boolean useBasePath, ExcelWorkBook excelWorkBook){
        this.excelWorkBookWriter.writeExcelWorkBook(filePath, useBasePath, excelWorkBook);
    }

    public ExcelWorkBook joinWorkBooksUsingKey(ExcelWorkBook workBook1, ExcelWorkBook workBook2){
        return this.excelWorkSheetMerger.joinWorkBooksUsingKey(workBook1, workBook2);
    }

    public ExcelWorkBook joinWorkBooksUSingKeySingleSheet(ExcelWorkBook workBook1, ExcelWorkBook workBook2){
        return this.excelWorkSheetMerger.joinWorkBookUsingKeySingleSheet(workBook1, workBook2);
    }

    public ExcelWorkBook generateWorkBookReport(ExcelWorkBook workBook){
        return this.excelReportGenerator.generateWorkBookReport(workBook);
    }
}
