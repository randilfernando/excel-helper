package com.alternate.officetools.service;

import com.alternate.officetools.helper.FileHelper;
import com.alternate.officetools.model.ExcelWorkBook;
import com.alternate.officetools.model.ExcelWorkSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class ExcelWorkBookReader {

    public ExcelWorkBook readExcelWorkBook(String filePath, boolean useBasePath, int keyColumnIndex){
        try(FileInputStream fileInputStream = FileHelper.getInputStream(filePath, useBasePath)) {
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            return this.extractExcelWorkBook(workbook, keyColumnIndex);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    private ExcelWorkBook extractExcelWorkBook(XSSFWorkbook workbook, int keyColumnIndex){
        ExcelWorkBook excelWorkBook = new ExcelWorkBook();

        int numberOfSheets = workbook.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++){
            XSSFSheet sheet = workbook.getSheetAt(i);
            ExcelWorkSheet excelWorkSheet = this.extractExcelWorkSheet(sheet, keyColumnIndex);
            excelWorkBook.addWorkSheet(excelWorkSheet);
        }

        return excelWorkBook;
    }

    private ExcelWorkSheet extractExcelWorkSheet(XSSFSheet sheet, int keyColumnIndex){
        String sheetName = sheet.getSheetName();
        ExcelWorkSheet excelWorkSheet = new ExcelWorkSheet(sheetName);

        Iterator<Row> rowIterator = sheet.rowIterator();

        if (rowIterator.hasNext()){
            Row row = rowIterator.next();
            Object[] extractedRow = this.extractExcelRow(row);
            excelWorkSheet.setColumnNames(extractedRow);
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Object[] extractedRow = this.extractExcelRow(row);
            if (extractedRow.length > 0){
                excelWorkSheet.insertRow(extractedRow[keyColumnIndex].toString(), extractedRow);
            }
        }

        return excelWorkSheet;
    }

    private Object[] extractExcelRow(Row row){
        Iterator<Cell> cellIterator = row.cellIterator();
        ArrayList<Object> objectArray = new ArrayList<>();

        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();

            switch (cell.getCellTypeEnum()) {
                case NUMERIC:
                    objectArray.add(cell.getNumericCellValue());
                    break;
                case STRING:
                    objectArray.add(cell.getStringCellValue());
                    break;
            }
        }

        return objectArray.toArray();
    }

}
