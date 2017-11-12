package com.alternate.officetools.application;

import com.alternate.officetools.helper.FileHelper;
import com.alternate.officetools.model.ExcelWorkBook;
import com.alternate.officetools.model.ExcelWorkSheet;
import com.alternate.officetools.service.ExcelHelperEngine;
import com.alternate.officetools.service.ExcelWorkBookReader;
import com.alternate.officetools.service.ExcelWorkBookWriter;
import com.alternate.officetools.service.ExcelWorkSheetMerger;

import java.util.ArrayList;
import java.util.Scanner;

public class ExcelHelperConsole implements IExcelHelper {
    private ExcelHelperEngine excelHelperEngine;

    public ExcelHelperConsole(){
        this.excelHelperEngine = new ExcelHelperEngine();
    }

    public void startApplication() {
        this.showMainMenu();
    }

    private void showMainMenu(){
        Scanner scanner = new Scanner(System.in);

        System.out.println("Excel sheet merger");
        System.out.println("==========================================================================================");

        System.out.print("Enter input file1 name: ");
        String inputFile1 = scanner.nextLine();

        System.out.print("Enter key column indexes: ");
        String[] keyColumns1 = scanner.nextLine().split(",");

        System.out.print("Enter grouping column index: ");
        String groupingColumn1 = scanner.nextLine();

        System.out.print("Enter input file2 name: ");
        String inputFile2 = scanner.nextLine();

        System.out.print("Enter key column indexes: ");
        String[] keyColumns2 = scanner.nextLine().split(",");

        System.out.print("Enter grouping column index: ");
        String groupingColumn2 = scanner.nextLine();

        System.out.print("Enter result file name: ");
        String outputFile = scanner.nextLine();

        mergeWorkbookSheets(inputFile1, inputFile2, outputFile, true, keyColumns1, keyColumns2, groupingColumn1, groupingColumn2);

        System.out.println("==========================================================================================");
        System.out.println("End");
    }

    private void mergeWorkbookSheets(String inputFile1, String inputFile2, String outputFile, boolean useBasePath,
                                     String[] keyColumns1, String[] keyColumns2, String groupingColumn1, String groupingColumn2){
        int[] keyColumnsReal1 = new int[keyColumns1.length];
        int groupingColumnReal1 = Integer.parseInt(groupingColumn1);

        int[] keyColumnsReal2 = new int[keyColumns2.length];
        int groupingColumnReal2 = Integer.parseInt(groupingColumn2);

        for (int i = 0; i < keyColumns1.length; i++){
            keyColumnsReal1[i] = Integer.parseInt(keyColumns1[i]);
        }

        for (int i = 0; i < keyColumns2.length; i++){
            keyColumnsReal2[i] = Integer.parseInt(keyColumns2[i]);
        }

        ExcelWorkBook excelWorkBook1 = this.excelHelperEngine.readExcelWorkBook(inputFile1, useBasePath, keyColumnsReal1, groupingColumnReal1);
        ExcelWorkBook excelWorkBook2 = this.excelHelperEngine.readExcelWorkBook(inputFile2, useBasePath, keyColumnsReal2, groupingColumnReal2);

        ExcelWorkBook resultExcelWorkBook = this.excelHelperEngine.joinWorkBooksUsingKey(excelWorkBook1, excelWorkBook2);

        this.excelHelperEngine.writeExcelWorkBook(outputFile, true, resultExcelWorkBook);

        System.out.println("");
    }
}
