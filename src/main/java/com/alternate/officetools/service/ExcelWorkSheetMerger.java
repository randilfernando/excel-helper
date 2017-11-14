package com.alternate.officetools.service;

import com.alternate.officetools.model.ExcelWorkBook;
import com.alternate.officetools.model.ExcelWorkSheet;
import org.apache.commons.lang3.ArrayUtils;

public class ExcelWorkSheetMerger {
    public ExcelWorkBook joinWorkBooksUsingKey(ExcelWorkBook workBook1, ExcelWorkBook workBook2) {
        ExcelWorkBook resultWorkBook = new ExcelWorkBook();

        ExcelWorkSheet excelWorkSheet1 = this.joinDataUsingKey(workBook1.getExcelWorkSheets().get(0), workBook2.getExcelWorkSheets().get(0));
        ExcelWorkSheet excelWorkSheet2 = this.joinDuplicateEntriesUsingKey(workBook1.getExcelWorkSheets().get(0), workBook2.getExcelWorkSheets().get(0));

        resultWorkBook.addWorkSheet(excelWorkSheet1);
        resultWorkBook.addWorkSheet(excelWorkSheet2);

        return resultWorkBook;
    }

    public ExcelWorkBook joinWorkBookUsingKeySingleSheet(ExcelWorkBook workBook1, ExcelWorkBook workBook2) {
        ExcelWorkBook resultWorkBook = new ExcelWorkBook();

        ExcelWorkSheet excelWorkSheet1 = this.joindDataUsingKeySingleSheet(workBook1.getExcelWorkSheets().get(0), workBook2.getExcelWorkSheets().get(0));
        ExcelWorkSheet excelWorkSheet2 = this.joinDuplicateEntriesUsingKeySingleSheet(workBook1.getExcelWorkSheets().get(0), workBook2.getExcelWorkSheets().get(0));

        resultWorkBook.addWorkSheet(excelWorkSheet1);
        resultWorkBook.addWorkSheet(excelWorkSheet2);

        return resultWorkBook;
    }

    private ExcelWorkSheet joindDataUsingKeySingleSheet(ExcelWorkSheet workSheet1, ExcelWorkSheet workSheet2) {
        ExcelWorkSheet resultWorkSheet = new ExcelWorkSheet();
        Object[] columns = ArrayUtils.addAll(workSheet1.getColumnNames(), workSheet2.getColumnNames());
        resultWorkSheet.setColumnNames(columns);

        for (String key : workSheet1.getData().keySet()) {
            if (workSheet2.isInDuplicateEntries(key)){
                Object[] firstRow = ArrayUtils.addAll(workSheet1.getData().get(key), (Object[]) workSheet2.getDuplicateEntries().get(key)[0]);
                resultWorkSheet.insertRow(firstRow, false);

                int length = workSheet2.getDuplicateEntries().get(key).length;

                for (int i = 1; i < length; i++){
                    Object[] resultRow = ArrayUtils.addAll(new Object[workSheet1.getColumnNames().length], (Object[]) workSheet2.getDuplicateEntries().get(key)[i]);
                    resultWorkSheet.insertRow(resultRow, false);
                }
            }

            Object[] resultRow = ArrayUtils.addAll(workSheet1.getData().get(key), workSheet2.getData().get(key));
            resultWorkSheet.insertRow(resultRow, false);
        }

        return resultWorkSheet;
    }

    private ExcelWorkSheet joinDuplicateEntriesUsingKeySingleSheet(ExcelWorkSheet workSheet1, ExcelWorkSheet workSheet2) {
        ExcelWorkSheet resultWorkSheet = new ExcelWorkSheet();
        Object[] columns = ArrayUtils.addAll(workSheet1.getColumnNames(), workSheet2.getColumnNames());
        resultWorkSheet.setColumnNames(columns);

        int workSheet1Length = workSheet1.getColumnNames().length;

        for (String key : workSheet1.getDuplicateEntries().keySet()) {
            for (Object row : workSheet1.getDuplicateEntries().get(key)) {
                resultWorkSheet.insertRow((Object[]) row, false);
            }

            if (workSheet2.isInData(key)) {
                Object[] resultRow = ArrayUtils.addAll(new Object[workSheet1Length], workSheet2.getData().get(key));
                resultWorkSheet.insertRow(resultRow, false);
            } else if (workSheet2.isInDuplicateEntries(key)) {
                for (Object row : workSheet2.getDuplicateEntries().get(key)) {
                    Object[] resultRow = ArrayUtils.addAll(new Object[workSheet1Length], (Object[]) row);
                    resultWorkSheet.insertRow(resultRow, false);
                }
            }
        }

        return resultWorkSheet;
    }

    private ExcelWorkSheet joinDataUsingKey(ExcelWorkSheet workSheet1, ExcelWorkSheet workSheet2) {
        ExcelWorkSheet resultWorkSheet = new ExcelWorkSheet();
        Object[] columns = ArrayUtils.addAll(workSheet1.getColumnNames(), workSheet2.getColumnNames());
        resultWorkSheet.setColumnNames(columns);

        for (String key : workSheet1.getData().keySet()) {
            if (workSheet2.isInDuplicateEntries(key)) continue;

            Object[] resultRow = ArrayUtils.addAll(workSheet1.getData().get(key), workSheet2.getData().get(key));
            resultWorkSheet.insertRow(resultRow, false);
        }

        return resultWorkSheet;
    }

    private ExcelWorkSheet joinDuplicateEntriesUsingKey(ExcelWorkSheet workSheet1, ExcelWorkSheet workSheet2) {
        ExcelWorkSheet resultWorkSheet = new ExcelWorkSheet();
        Object[] columns = ArrayUtils.addAll(workSheet1.getColumnNames(), workSheet2.getColumnNames());
        resultWorkSheet.setColumnNames(columns);

        int workSheet1Length = workSheet1.getColumnNames().length;

        for (String key : workSheet1.getDuplicateEntries().keySet()) {
            for (Object row : workSheet1.getDuplicateEntries().get(key)) {
                resultWorkSheet.insertRow((Object[]) row, false);
            }

            if (workSheet2.isInData(key)) {
                Object[] resultRow = ArrayUtils.addAll(new Object[workSheet1Length], workSheet2.getData().get(key));
                resultWorkSheet.insertRow(resultRow, false);
            } else if (workSheet2.isInDuplicateEntries(key)) {
                for (Object row : workSheet2.getDuplicateEntries().get(key)) {
                    Object[] resultRow = ArrayUtils.addAll(new Object[workSheet1Length], (Object[]) row);
                    resultWorkSheet.insertRow(resultRow, false);
                }
            }
        }

        for (String key : workSheet2.getDuplicateEntries().keySet()) {
            if (workSheet1.isInData(key)) {
                resultWorkSheet.insertRow(workSheet1.getData().get(key), false);
                for (Object row : workSheet2.getDuplicateEntries().get(key)) {
                    Object[] resultRow = ArrayUtils.addAll(new Object[workSheet1Length], (Object[]) row);
                    resultWorkSheet.insertRow(resultRow, false);
                }
            }
        }

        return resultWorkSheet;
    }
}
