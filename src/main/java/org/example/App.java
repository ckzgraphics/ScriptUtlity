package org.example;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.sample.BaseWebDriver;

import java.util.ArrayList;
import java.util.List;

public class App 
{
    public static void main( String[] args ) {
        String propWorkbookName = System.getProperty("book");
        String url = System.getProperty("url");
        new BaseWebDriver(propWorkbookName, url);
    }
}
