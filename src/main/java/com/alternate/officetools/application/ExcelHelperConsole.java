package com.alternate.officetools.application;

import com.alternate.officetools.model.ExcelWorkBook;
import com.alternate.officetools.model.ExcelWorkSheet;
import com.alternate.officetools.service.ExcelWorkBookReader;
import com.alternate.officetools.service.ExcelWorkBookWriter;
import com.alternate.officetools.service.ExcelWorkSheetMerger;

import java.util.Scanner;

public class ExcelHelperConsole implements IExcelHelper {
    private ExcelWorkBookReader excelWorkBookReader;
    private ExcelWorkBookWriter excelWorkBookWriter;
    private ExcelWorkSheetMerger excelWorkSheetMerger;

    public ExcelHelperConsole(){
        this.excelWorkBookReader = new ExcelWorkBookReader();
        this.excelWorkBookWriter = new ExcelWorkBookWriter();
        this.excelWorkSheetMerger = new ExcelWorkSheetMerger();
    }

    public void startApplication() {
        this.showMainMenu();
    }

    private void showMainMenu(){
        Scanner scanner = new Scanner(System.in);

        System.out.println("Excel sheet merger");
        System.out.println("==========================================================================================");

        System.out.print("Enter input file name: ");
        String inputFile = scanner.nextLine();

        System.out.print("Enter output file name: ");
        String outputFile = scanner.nextLine();

        System.out.printf("Enter key column index: ");
        int keyColumnIndex = scanner.nextInt();

        mergeWorkbookSheets(inputFile, outputFile, keyColumnIndex);

        System.out.println("==========================================================================================");
        System.out.println("End");
    }

    private void mergeWorkbookSheets(String inputFile, String outputFile, int keyColumnIndex){
        ExcelWorkBook excelWorkBook = this.excelWorkBookReader.readExcelWorkBook(inputFile, true, keyColumnIndex);

        ExcelWorkSheet excelWorkSheet = this.excelWorkSheetMerger.leftJoinWorkSheetsUsingKeyColumn(
                excelWorkBook.getExcelWorkSheets().get(0),
                excelWorkBook.getExcelWorkSheets().get(1),
                keyColumnIndex
        );

        ExcelWorkBook resultWorkBook = new ExcelWorkBook("result.xlsx");
        resultWorkBook.addWorkSheet(excelWorkSheet);

        this.excelWorkBookWriter.writeExcelWorkBook(outputFile, true, resultWorkBook);
    }
}
