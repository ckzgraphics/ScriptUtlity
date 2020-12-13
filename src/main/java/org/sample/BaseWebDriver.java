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

    public BaseWebDriver() {
        exe = new ExcelHandler(Base.PATH_DATA, Base.WORKBOOK_NAME);
        int sheetCount = exe.getNumberOfSheets();
        System.out.println("Total No of Sheets found: " + sheetCount);

        if (sheetCount > 0) {
            symbol = exe.getSheetNameList();
            initDriver();
        } else {
            System.out.println("Terminating utility");
        }
    }

    private void initDriver() {
        options = new ChromeOptions();
        options.addArguments("--incognito", "--disable-blink-features=AutomationControlled");
        WebDriverManager.chromedriver().setup();

        driver = new ChromeDriver(options);
        driver.manage().window().maximize();
        js = (JavascriptExecutor) driver;

        js.executeScript("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})");
        js.executeScript("'Network.setUserAgentOverride', {\"userAgent\": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.53 Safari/537.36'}");
    }

    public void start() {
        try {
            for (String symbolName : symbol) {

                newDateData = new ArrayList<String>();
                System.out.println("\nExecuting for: " + symbolName);
//                exe.printSheetData(symbolName);

                driver.get(Base.BASE_URL.concat(symbolName));

                sleep(3000);
                scroll(500);

//                System.out.println("\n" + driver.getPageSource());

                driver.findElement(By.xpath(".//h2[contains(.,'Historical Data')]")).click();
                sleep(3000);

                driver.findElement(By.xpath("(.//ul[@class='dayslisting'])[1]/li[contains(.,'1M')]/a")).click();
                sleep(3000);

                List<WebElement> tableHeaders = driver.findElements(By.xpath(".//table[@id='equityHistoricalTable']/thead/tr[1]/th"));
                List<WebElement> tableData = driver.findElements(By.xpath(".//table[@id='equityHistoricalTable']/tbody/tr[1]/td"));
                for (int i = 0; i < tableHeaders.size(); i++) {

                    if (i == 1) {
                        continue;
                    }

                    String newDateDataStr = tableData.get(i).getText();
                    System.out.println("=> " + tableHeaders.get(i).getText() + ": " + newDateDataStr);
                    newDateData.add(newDateDataStr);
                }

                exe.addRowData(symbolName, Base.PATH_DATA, Base.WORKBOOK_NAME, newDateData);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
//            if (Objects.nonNull(exe)) {
//
//            }
            if (Objects.nonNull(driver)) {
                driver.quit();
                exe.closeWorkbook();
            }
        }
    }

    private void sleep(int time) {
        try {
            Thread.sleep(time);
        } catch (Exception e) {
        }
    }

    private void scroll(int pixels) {
        js.executeScript("window.scrollBy(0," + pixels + ")");
    }
}
