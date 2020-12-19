package org.example;

import static org.junit.Assert.assertTrue;

import org.sample.BaseWebDriver;
import org.testng.annotations.Test;

public class AppTest 
{
    @Test
    public void initUtil(){
        String propWorkbookName = System.getProperty("book");
        String url = System.getProperty("url");
        new BaseWebDriver(propWorkbookName, url);
    }
}
