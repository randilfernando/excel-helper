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

    public static int[] conventIntoIntegerArray(String[] arrayIn){
        int[] arrayOut = new int[arrayIn.length];

        for (int i = 0; i < arrayIn.length; i++){
            arrayOut[i] = Integer.parseInt(arrayIn[i]);
        }

        return arrayOut;
    }

    public static String generateReportCellName(String prefix, Object[] row, int[] columns){
        StringBuilder name = new StringBuilder(prefix);

        for (int i: columns){
            name.append("_").append(row[i]);
        }

        return name.toString();
    }

}
