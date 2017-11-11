package com.alternate.officetools.helper;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

public final class FileHelper {
    private static final String basePath = "/Users/randilfernando/Documents/";

    private FileHelper(){}

    public static FileInputStream getInputStream(String filePath, boolean useBasePath) throws FileNotFoundException {
        if (useBasePath) filePath = basePath + filePath;
        return new FileInputStream(filePath);
    }

    public static FileOutputStream getOutPutStream(String filePath, boolean useBasePath) throws FileNotFoundException {
        if (useBasePath) filePath = basePath + filePath;
        return new FileOutputStream(filePath);
    }
}
