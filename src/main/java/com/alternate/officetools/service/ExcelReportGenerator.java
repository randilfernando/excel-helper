package com.alternate.officetools.service;

import com.alternate.officetools.helper.ExcelHelper;
import com.alternate.officetools.model.ExcelWorkBook;
import com.alternate.officetools.model.ExcelWorkSheet;

public class ExcelReportGenerator {
    public ExcelWorkBook generateWorkBookReport(ExcelWorkBook workBook){
        ExcelWorkBook resultWorkBook = new ExcelWorkBook();

        for (ExcelWorkSheet excelWorkSheet: workBook.getExcelWorkSheets()){
            resultWorkBook.addWorkSheet(this.generateWorkSheetReport(excelWorkSheet));
        }

        return resultWorkBook;
    }


    private ExcelWorkSheet generateWorkSheetReport(ExcelWorkSheet workSheet){
        ExcelWorkSheet resultWorkSheet = new ExcelWorkSheet();

        String groupsColumnName = ExcelHelper.generateReportCellName("Groups", workSheet.getColumnNames(), workSheet.getGroupingColumn());
        String keyColumnName = ExcelHelper.generateReportCellName("Keys", workSheet.getColumnNames(), workSheet.getKeyColumns());

        int groups = workSheet.getGroups().size();
        int keysWithoutDuplicates = workSheet.getData().size();
        int keysWithDuplicates = workSheet.getDuplicateEntries().size();
        int plainEntries = 0;
        int duplicateEntries = 0;

        plainEntries = workSheet.getData().size();

        for (String key: workSheet.getDuplicateEntries().keySet()){
            duplicateEntries += workSheet.getDuplicateEntries().get(key).length;
        }

        resultWorkSheet.setColumnNames(new Object[]{workSheet.getName()});
        resultWorkSheet.insertRow(new Object[]{groupsColumnName, groups}, false);
        resultWorkSheet.insertRow(new Object[]{keyColumnName + "(Without duplicates)", keysWithoutDuplicates}, false);
        resultWorkSheet.insertRow(new Object[]{keyColumnName + "(With duplicates)", keysWithDuplicates}, false);

        resultWorkSheet.insertRow(new Object[0], false);

        resultWorkSheet.insertRow(new Object[]{"Plain entries", plainEntries}, false);
        resultWorkSheet.insertRow(new Object[]{"Duplicate entries", duplicateEntries}, false);
        resultWorkSheet.insertRow(new Object[]{"Total entries", plainEntries + duplicateEntries}, false);

        return resultWorkSheet;
    }
}
