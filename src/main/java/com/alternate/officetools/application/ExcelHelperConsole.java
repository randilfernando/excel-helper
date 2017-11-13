package com.alternate.officetools.application;

import com.alternate.officetools.helper.ExcelHelper;
import com.alternate.officetools.model.ExcelWorkBook;
import com.alternate.officetools.service.ExcelHelperEngine;

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

        System.out.println("Excel Helper");
        System.out.println("==========================================================================================");

        int choice = 0;

        while (choice != 4){
            this.printMainMenu();

            System.out.print("Enter your choice: ");
            choice = Integer.parseInt(scanner.nextLine());

            switch (choice){
                case 1: mergeWorkbookSheets(false); break;
                case 2: mergeWorkbookSheets(true); break;
                case 3: generateWorkBookReport(); break;
                default:
                    System.out.println("Invalid input");
            }
        }

        System.out.println("==========================================================================================");
        System.out.println("End");
    }

    private void printMainMenu(){
        System.out.println("==========================================================================================");
        System.out.println("(1) Merge excel workbooks (Duplicates in seperate sheet)");
        System.out.println("(2) Merge excel workbooks (Duplicates in same sheet)");
        System.out.println("(3) Generate workbook report");
        System.out.println("(4) Exit");
        System.out.println("==========================================================================================");
    }

    private void generateWorkBookReport(){
        Scanner scanner = new Scanner(System.in);

        System.out.print("Enter input file name: ");
        String inputFile1 = scanner.nextLine();

        System.out.print("Enter key column indexes: ");
        String[] keyColumns1 = scanner.nextLine().split(",");

        System.out.print("Enter grouping column indexes: ");
        String[] groupingColumns1 = scanner.nextLine().split(",");

        System.out.print("Enter result file name: ");
        String outputFile = scanner.nextLine();

        int[] keyColumnsReal1 = ExcelHelper.conventIntoIntegerArray(keyColumns1);
        int[] groupingColumnsReal1 = ExcelHelper.conventIntoIntegerArray(groupingColumns1);

        System.out.println("Reading " + inputFile1);
        ExcelWorkBook excelWorkBook1 = this.excelHelperEngine.readExcelWorkBook(inputFile1, true, keyColumnsReal1, groupingColumnsReal1);
        System.out.println("Done");

        System.out.println("Generating report");
        ExcelWorkBook resultExcelWorkBook = this.excelHelperEngine.generateWorkBookReport(excelWorkBook1);
        System.out.println("Done");

        System.out.println("Writing " + outputFile);
        this.excelHelperEngine.writeExcelWorkBook(outputFile, true, resultExcelWorkBook);
        System.out.println("Done");

        System.out.println("Success");
    }

    private void mergeWorkbookSheets(boolean sameSheet){
        Scanner scanner = new Scanner(System.in);

        System.out.print("Enter input file1 name: ");
        String inputFile1 = scanner.nextLine();

        System.out.print("Enter key column indexes: ");
        String[] keyColumns1 = scanner.nextLine().split(",");

        System.out.print("Enter grouping column indexes: ");
        String[] groupingColumns1 = scanner.nextLine().split(",");

        System.out.print("Enter input file2 name: ");
        String inputFile2 = scanner.nextLine();

        System.out.print("Enter key column indexes: ");
        String[] keyColumns2 = scanner.nextLine().split(",");

        System.out.print("Enter grouping column index: ");
        String[] groupingColumns2 = scanner.nextLine().split(",");

        System.out.print("Enter result file name: ");
        String outputFile = scanner.nextLine();

        int[] keyColumnsReal1 = ExcelHelper.conventIntoIntegerArray(keyColumns1);
        int[] groupingColumnsReal1 = ExcelHelper.conventIntoIntegerArray(groupingColumns1);

        int[] keyColumnsReal2 = ExcelHelper.conventIntoIntegerArray(keyColumns2);
        int[] groupingColumnsReal2 = ExcelHelper.conventIntoIntegerArray(groupingColumns2);

        System.out.println("Reading " + inputFile1);
        ExcelWorkBook excelWorkBook1 = this.excelHelperEngine.readExcelWorkBook(inputFile1, true, keyColumnsReal1, groupingColumnsReal1);
        System.out.println("Done");

        System.out.println("Reading " + inputFile2);
        ExcelWorkBook excelWorkBook2 = this.excelHelperEngine.readExcelWorkBook(inputFile2, true, keyColumnsReal2, groupingColumnsReal2);
        System.out.println("Done");

        System.out.println("Merging " + inputFile1 + " + " + inputFile2);

        ExcelWorkBook resultExcelWorkBook;

        if (sameSheet){
            resultExcelWorkBook = this.excelHelperEngine.joinWorkBooksUSingKeySingleSheet(excelWorkBook1, excelWorkBook2);
        }else{
            resultExcelWorkBook = this.excelHelperEngine.joinWorkBooksUsingKey(excelWorkBook1, excelWorkBook2);
        }

        System.out.println("Done");

        System.out.println("Writing " + outputFile);
        this.excelHelperEngine.writeExcelWorkBook(outputFile, true, resultExcelWorkBook);
        System.out.println("Done");

        System.out.println("Success");
    }
}
