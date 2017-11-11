package com.alternate.officetools.service;

import com.alternate.officetools.helper.FileHelper;
import com.alternate.officetools.model.ExcelWorkBook;
import com.alternate.officetools.model.ExcelWorkSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWorkBookWriter {
    public void writeExcelWorkBook(String filePath, boolean useBasePath, ExcelWorkBook excelWorkBook) {
        try (FileOutputStream fileOutputStream = FileHelper.getOutPutStream(filePath, useBasePath)) {
            XSSFWorkbook workbook = this.createWorkBook(excelWorkBook);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private XSSFWorkbook createWorkBook(ExcelWorkBook excelWorkBook) {
        XSSFWorkbook workbook = new XSSFWorkbook();

        for (ExcelWorkSheet excelWorkSheet : excelWorkBook.getExcelWorkSheets()) {
            XSSFSheet sheet = workbook.createSheet(excelWorkSheet.getName());
            this.createWorkSheet(excelWorkSheet, sheet);
        }

        return workbook;
    }

    private void createWorkSheet(ExcelWorkSheet excelWorkSheet, XSSFSheet sheet) {
        int rowNumber = 0;

        Row columns = sheet.createRow(rowNumber++);

        this.createRow(excelWorkSheet.getColumnNames(), columns);

        for (String key : excelWorkSheet.getData().keySet()) {
            Row row = sheet.createRow(rowNumber++);

            this.createRow(excelWorkSheet.getData().get(key), row);
        }
    }

    private void createRow(Object[] objArr, Row row) {
        int cellNumber = 0;

        for (Object obj : objArr) {
            Cell cell = row.createCell(cellNumber++);
            if (obj instanceof String) {
                cell.setCellValue((String) obj);
            } else if (obj instanceof Integer) {
                cell.setCellValue((Integer) obj);
            } else if (obj instanceof Double) {
                cell.setCellValue((Double) obj);
            } else {
                cell.setCellValue("");
            }
        }
    }
}
