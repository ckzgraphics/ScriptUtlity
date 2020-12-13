package org.base.global;

import java.io.File;

public class Base {

    public static final String BASE_PATH = System.getProperty("user.dir").concat(File.separator);
    public static final String PATH_DATA = BASE_PATH.concat("data").concat(File.separator);
    public static final String BASE_URL = "";
    public static final String WORKBOOK_NAME = "POC.xlsx";

}
