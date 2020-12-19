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

public class BaseWebDriver {

    private List<String> symbol = new ArrayList<String>();
    private List<String> newDateData = null;

    private WebDriver driver = null;
    private ChromeOptions options = null;
    private JavascriptExecutor js = null;

    private ExcelHandler exe = null;

    public BaseWebDriver(String workbookName, String url) {
        if(Objects.isNull(workbookName) || workbookName.isEmpty()){
            System.out.println("--- NO ENV VAR FOUND | USING DEFAULT FILE NAME ---");
            workbookName = Base.WORKBOOK_NAME;
        }
        exe = new ExcelHandler(Base.PATH_DATA, workbookName);
        int sheetCount = exe.getNumberOfSheets();
        System.out.println("Total No of Sheets found => " + sheetCount);

        if (sheetCount > 0) {
            symbol = exe.getSheetNameList();
            if(initDriver())
            {
                initProcess(url);
            }
        } else {
            System.out.println("Terminating utility: Total Sheets found: " + sheetCount);
        }
    }

    private boolean initDriver() {
        try{
            options = new ChromeOptions();
            options.addArguments("--incognito", "--disable-blink-features=AutomationControlled");
            WebDriverManager.chromedriver().setup();

            driver = new ChromeDriver(options);
            driver.manage().window().maximize();
            js = (JavascriptExecutor) driver;

            js.executeScript("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})");
            js.executeScript("'Network.setUserAgentOverride', {\"userAgent\": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.53 Safari/537.36'}");
            return true;
        } catch (Exception e){
            System.out.println("Error while initiating browser");
            e.printStackTrace();
            return false;
        }
    }

    private void initProcess(String url) {
        try {
            for (String symbolName : symbol) {

                if(symbolName.equalsIgnoreCase("Summary") || symbolName.equalsIgnoreCase("summary")){
                    continue;
                }

                newDateData = new ArrayList<String>();
                System.out.println("\n=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
                        +"\nExecuting for : " + symbolName);

                boolean isDataParseComplete = false;

                try {
                    // OPEN BROWSER
                    driver.get(url.concat(symbolName));

                    // WAIT FOR PAGE TO LOAD
                    sleep(3000);
                    // SCROLL TILL TABLE
                    scrollVertical(500);

                    driver.findElement(By.xpath(".//h2[contains(.,'Historical Data')]")).click();
                    sleep(3000);
                    driver.findElement(By.xpath("(.//ul[@class='dayslisting'])[1]/li[contains(.,'1M')]/a")).click();
                    sleep(3000);

                    List<WebElement> tableHeaders = driver.findElements(By.xpath(".//table[@id='equityHistoricalTable']/thead/tr[1]/th"));
                    List<WebElement> tableData = driver.findElements(By.xpath(".//table[@id='equityHistoricalTable']/tbody/tr[1]/td"));

                    String newDateDataStr = "";
                    String newDateDataStrPrint = "";
                    String newColHeaderStrPrint = "";

                    for (int i = 0; i < tableHeaders.size(); i++) {

                        // IGNORE COLUMN - SERIES
                        if (i == 1) {
                            continue;
                        }
                        newColHeaderStrPrint += tableHeaders.get(i).getText() + "|";
                        newDateDataStr = tableData.get(i).getText();
                        newDateDataStrPrint += newDateDataStr + "|";
                        // CREATE CELLs AND SET VALUES
                        newDateData.add(newDateDataStr);
                    }

                    System.out.println("\nHEADER: " + newColHeaderStrPrint);
                    System.out.println("DATA: " + newDateDataStrPrint + "\n");

                    isDataParseComplete = true;

                } catch(Exception e){
                    isDataParseComplete = false;
                    System.out.println("Error while reading data for: " + symbolName);
                    e.printStackTrace();
                } finally {
                    if(isDataParseComplete){
                        // ADD DATA TO SHEET
                        exe.addRowData(symbolName, newDateData);
                    }
                }
            } // FOR END - SYMBOL

            // SAVE WORKBOOK TO DISC
            exe.saveFile(Base.PATH_DATA, Base.WORKBOOK_NAME);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (Objects.nonNull(exe))
                exe.closeWorkbook();
            if (Objects.nonNull(driver))
                driver.quit();
        }
        System.out.println("\n=-=-=-=-=-=-=-=E=-=-=-N-=-=-=D=-=-=-=-=-=-=-=\n");
    }

    private void sleep(int time) {
        try { Thread.sleep(time); } catch (Exception e) { System.out.println(e.getMessage());}
    }

    private void scrollVertical(int pixels) {
        js.executeScript("window.scrollBy(0," + pixels + ")");
    }
}
