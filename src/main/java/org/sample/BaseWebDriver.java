package org.sample;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.base.global.Base;
import org.base.handler.ExcelHandler;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;

/**
 * Hello world!
 *
 */
public class BaseWebDriver
{

    WebDriver driver = null;
    ChromeOptions options = new ChromeOptions();
    String autURL = "https://www.nseindia.com/get-quotes/equity?symbol=";
    List<String> symbol = new ArrayList<String>();
    List<String> newDateData = null;
    JavascriptExecutor js = null;

    public void start()
    {
        ExcelHandler exe = new ExcelHandler(Base.PATH_DATA, Base.WORKBOOK_NAME);

        System.out.println( "Starting... No of Sheets = " + exe.getNumberOfSheets());
        symbol = exe.getSheetNameList();

        options.addArguments("--incognito", "--disable-blink-features=AutomationControlled");
        WebDriverManager.chromedriver().setup();

        driver = new ChromeDriver(options);
        driver.manage().window().maximize();
        js = (JavascriptExecutor) driver;

        try {
            js.executeScript("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})");

            for (String symbolName : symbol){

                newDateData = new ArrayList<String>();
                System.out.println("\nExecuting for : " + symbolName);
                exe.printSheetData(symbolName);

                driver.get(autURL.concat(symbolName));
                sleep(3000);
                scroll(500);

                driver.findElement(By.xpath(".//h2[contains(.,'Historical Data')]")).click();
                sleep(3000);

                driver.findElement(By.xpath("(.//ul[@class='dayslisting'])[1]/li[contains(.,'1M')]/a")).click();
                sleep(3000);


                List<WebElement> tableHeaders = driver.findElements(By.xpath(".//table[@id='equityHistoricalTable']/thead/tr[1]/th"));
                List<WebElement> tableData = driver.findElements(By.xpath(".//table[@id='equityHistoricalTable']/tbody/tr[1]/td"));
                for(int i=0; i<tableHeaders.size(); i++) {

                    String newDateDataStr = tableData.get(i).getText();
                    System.out.println("=> " + tableHeaders.get(i).getText()
                            + " : " + newDateDataStr);
                    newDateData.add(newDateDataStr);
                }

                exe.addRowData(symbolName, Base.PATH_DATA, Base.WORKBOOK_NAME, newDateData);
            }
        } catch(Exception e) {e.printStackTrace();}
        finally {
            if(Objects.nonNull(driver)) {
                exe.closeWorkbook();
                driver.quit();
            }
        }
    }

    public void sleep(int time){
        try { Thread.sleep(time); } catch(Exception e){}
    }

    public void scroll(int pixels){
        js.executeScript("window.scrollBy(0," + pixels + ")");
    }
}
