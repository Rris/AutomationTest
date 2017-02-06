package com.fsoft.selenium.testng;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import junit.framework.TestCase;

/**
 * Unit test for simple App.
 */
public class AppTest extends TestCase {
    private static WebDriver webDriver;
    final static int MAX = 50;
    final static Logger logger = Logger.getLogger(App.class);

    public void setupBrowser(String inputValue) {
        switch (inputValue) {
        case "Chrome":
            System.setProperty("webdriver.chrome.driver", "D:/chromedriver.exe");
            ChromeOptions options = new ChromeOptions();
            options.addArguments("--start-maximized");
            webDriver = new ChromeDriver(options);
            break;
        case "IE":
            System.setProperty("webdriver.ie.driver", "D:/MicrosoftWebDriver.exe");
            webDriver = new InternetExplorerDriver();
            break;
        case "Firefox":
            System.setProperty("webdriver.gecko.driver", "D:/geckodriver.exe");
            webDriver = new FirefoxDriver();
            break;
        default:
            break;
        }
    }

    public void loadPage(String inputValue) {
        webDriver.get(inputValue);
    }
    
    public void switchFrame(String inputValue) {
        switch (inputValue) {
        case "default":
            webDriver.switchTo().defaultContent();
            break;
        case "contentFrame":
            webDriver.switchTo().parentFrame();
        default:
            break;
        }
    }

    public void enterInput(WebElement webElement, String inputValue) {
        webElement.sendKeys(inputValue);
    }
    
    public boolean verify_element_present(){
        return true;
    }

    public WebElement findElement(String locatorId, String locatorString) {
        WebElement webElement = null;
        switch (locatorId) {
        case "id":
            webElement = webDriver.findElement(By.id(locatorString));
            break;
        case "name":
            webElement = webDriver.findElement(By.name(locatorString));
            break;
        case "className":
            webElement = webDriver.findElement(By.className(locatorString));
            break;
        case "cssSelector":
            webElement = webDriver.findElement(By.cssSelector(locatorString));
            break;
        case "linkText":
            webElement = webDriver.findElement(By.linkText(locatorString));
            break;
        case "partialLinkText":
            webElement = webDriver.findElement(By.partialLinkText(locatorString));
            break;
        case "tagName":
            webElement = webDriver.findElement(By.tagName(locatorString));
            break;
        case "xpath":
            webElement = webDriver.findElement(By.xpath(locatorString));
            break;
        default:
            break;
        }
        return webElement;
    }

    public void clickElement(WebElement webElement) {
        webElement.click();
    }

    public void dragAndDrop(String locatorId, String locatorString, String inputValue) {
        WebElement fromWebElement1 = findElement(locatorId, locatorString);
        WebElement toWebElement2 = findElement(locatorId, inputValue);
        (new Actions(webDriver)).dragAndDrop(fromWebElement1, toWebElement2).perform();
    }

    public void screenShot(String inputValue) {
        File scrFile = ((TakesScreenshot) webDriver).getScreenshotAs(OutputType.FILE);
        try {
            FileUtils.copyFile(scrFile, new File("/SelScripts/TestAuto_Framework/screenshots/" + inputValue + ".png"));
        } catch (IOException e) {
            logger.error("File not found!");
        }
    }

    public void waitTime(String inputValue) {
        int waitTime = Integer.parseInt(inputValue);
        webDriver.manage().timeouts().implicitlyWait(waitTime, TimeUnit.SECONDS);
    }

    public void startLoop(int loopTimes, int fromLine, int toLine) {
        String[][] storeData = (String[][]) prepareData();
        for (int i = 0; i < loopTimes; i++) {
            for (int j = fromLine - 2; j < toLine - 2; j++) {
                excuseKeyword(storeData[j]);
            }
        }
    }
    
    public void deselectDropDown(String locatorId, String locatorString) {
        Select select = new Select(findElement(locatorId, locatorString));
        select.deselectAll();
        //select.deselectByIndex(0);
    }
    
    public void closeDriver() {
        webDriver.close();
    }
    
    public void quitDriver() {
        webDriver.quit();
    }
    
    public void sendKey() {
        Actions builder = new Actions(webDriver);
        builder.keyDown(Keys.TAB).perform();
    }

    @Test(dataProvider = "data")
    public void excuseKeyword(String[] args) {
        String keyword = args[0].trim();
        String locatorId = args[1].trim();
        String locatorString = args[2].trim();
        String inputValue = args[3].trim();
        int numbersOfInputParam = Integer.parseInt(args[4]);
        int numbersOfOutputParam = Integer.parseInt(args[5]);
        String inputQuery = args[6].trim();
        String expected = args[7].trim();
        switch (keyword) {
        case "setup_browser":
            setupBrowser(inputValue);
            break;
        case "load_page":
            loadPage(inputValue);
            break;
        case "click_element":
            clickElement(findElement(locatorId, locatorString));
            break;
        case "enter_input":
            enterInput(findElement(locatorId, locatorString), inputValue);
            break;
        case "drag_and_drop":
            dragAndDrop(locatorId, locatorString, inputValue);
            break;
        case "screen_shot":
            screenShot(inputValue);
            break;
        case "wait_time":
            waitTime(inputValue);
            break;
        case "start_loop":
            int loopTimes = Integer.parseInt(args[8]);
            int fromLine = Integer.parseInt(args[9]);
            int toLine = Integer.parseInt(args[10]);
            startLoop(loopTimes, fromLine, toLine);
            break;
        default:
            break;
        }
    }

    @DataProvider(name = "data")
    public Object[][] prepareData() {
        Object[][] arrayObject = getDataFromExcelFile("D:/4.1.RITE20_ModuleName_UI_FunctionName_TestCases.xlsx",
                "TC01_UI");
        return arrayObject;
    }

    public String[][] getDataFromExcelFile(String filePath, String sheetName) {
        String[][] arrKeywords = null;
        int keyword = 0;
        int locatorId = 0;
        int locatorString = 0;
        int inputValue = 0;
        int numbersOfInputParam = 0;
        int numbersOfOutputParam = 0;
        int inputQuery = 0;
        int expected = 0;
        int totalRows = 0;
        try {
            FileInputStream fileInput = new FileInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(fileInput);
            XSSFSheet sheet = workbook.getSheet(sheetName);
            totalRows = sheet.getPhysicalNumberOfRows();
            int totalColumns = sheet.getRow(0).getPhysicalNumberOfCells();
            arrKeywords = new String[totalRows - 1][totalColumns];
            Iterator<Row> rows = sheet.rowIterator();
            while (rows.hasNext()) {
                XSSFRow row = (XSSFRow) rows.next();
                int rowNumber = row.getRowNum();
                for (int i = 0; i < totalColumns; i++) {
                    if (rowNumber == 0) {
                        switch (row.getCell(i).getStringCellValue()) {
                        case "Keyword":
                            keyword = i;
                            break;
                        case "Locator Id":
                            locatorId = i;
                            break;
                        case "Locate String":
                            locatorString = i;
                            break;
                        case "Input Value":
                            inputValue = i;
                            break;
                        case "#Input Params":
                            numbersOfInputParam = i;
                            break;
                        case "#Output Params":
                            numbersOfOutputParam = i;
                            break;
                        case "Input Query":
                            inputQuery = i;
                            break;
                        case "Expected":
                            expected = i;
                            break;
                        default:
                            break;
                        }
                    } else {
                        row.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                        arrKeywords[row.getRowNum() - 1][i] = row.getCell(i).getStringCellValue();
                    }
                }
            }
            workbook.close();
        } catch (Exception ioe) {
            ioe.printStackTrace();
        }

        String[][] keywords = new String[totalRows - 1][MAX];
        for (int i = 0; i < totalRows - 1; i++) {
            int pos = 0;
            keywords[i][pos++] = arrKeywords[i][keyword];
            keywords[i][pos++] = arrKeywords[i][locatorId];
            keywords[i][pos++] = arrKeywords[i][locatorString];
            keywords[i][pos++] = arrKeywords[i][inputValue];
            keywords[i][pos] = arrKeywords[i][numbersOfInputParam] == "" ? "0" : arrKeywords[i][numbersOfInputParam];
            int inputParams = Integer.parseInt(keywords[i][pos++]);
            keywords[i][pos] = arrKeywords[i][numbersOfOutputParam] == "" ? "0" : arrKeywords[i][numbersOfOutputParam];
            int outputParams = Integer.parseInt(keywords[i][pos++]);
            keywords[i][pos++] = arrKeywords[i][inputQuery];
            keywords[i][pos++] = arrKeywords[i][expected];
            for (int j = 0; j < inputParams; j++) {
                keywords[i][pos++] = arrKeywords[i][numbersOfInputParam + j + 1];
            }
            for (int j = 0; j < outputParams; j++) {
                keywords[i][pos++] = arrKeywords[i][numbersOfOutputParam + j + 1];
            }
        }
        return keywords;
    }
}
