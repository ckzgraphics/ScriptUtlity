package org.example;

import com.sun.codemodel.internal.JForEach;
import io.github.bonigarcia.wdm.WebDriverManager;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.sample.BaseWebDriver;

import java.util.ArrayList;
import java.util.List;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        BaseWebDriver b = new BaseWebDriver();
        b.start();

    }
}
