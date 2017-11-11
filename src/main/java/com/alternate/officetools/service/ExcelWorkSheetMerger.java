package com.alternate.officetools.service;

import com.alternate.officetools.model.ExcelWorkSheet;
import org.apache.commons.lang3.ArrayUtils;

public class ExcelWorkSheetMerger {
    public ExcelWorkSheet leftJoinWorkSheetsUsingKeyColumn(ExcelWorkSheet inputSheet1, ExcelWorkSheet inputSheet2, int keyColumnIndex) {
        ExcelWorkSheet resultWorkSheet = new ExcelWorkSheet();
        resultWorkSheet.setColumnNames(ArrayUtils.addAll(inputSheet1.getColumnNames(), ArrayUtils.remove(inputSheet2.getColumnNames(), keyColumnIndex)));

        for (String key : inputSheet1.getData().keySet()) {
            Object[] resultRow;

            if (inputSheet2.getData().containsKey(key)) {
                resultRow = ArrayUtils.addAll(inputSheet1.getData().get(key), ArrayUtils.remove(inputSheet2.getData().get(key), keyColumnIndex));
            } else {
                resultRow = inputSheet1.getData().get(key);
            }

            resultWorkSheet.insertRow(key, resultRow);
        }

        return resultWorkSheet;
    }
}
