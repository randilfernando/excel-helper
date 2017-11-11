package com.alternate.officetools;

import com.alternate.officetools.application.ExcelHelperConsole;
import com.alternate.officetools.application.IExcelHelper;

/**
 * Hello world!
 *
 */
public class App {
    private static final IExcelHelper excelHelperConsole = new ExcelHelperConsole();

    public static void main( String[] args ) {
        excelHelperConsole.startApplication();
    }
}
