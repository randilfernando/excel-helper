package com.alternate.officetools.helper;

import java.util.Date;

public final class ExcelHelper {

    private ExcelHelper(){}

    public static String generateNameUsingDate(String baseName){
        Date date = new Date();
        return baseName + "_" + date.toString();
    }

    public static String generateNameUsingWorkSheetNumber(String baseName, int number){
        return baseName + "_" + number;
    }

}
