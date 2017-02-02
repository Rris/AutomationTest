package com.fsoft.selenium.testng;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.commons.io.FileUtils;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import junit.framework.TestCase;

/**
 * Unit test for simple App.
 */
public class AppTest extends TestCase {
    private static WebDriver webDriver;
    final static int MAX = 100;
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

    public WebElement findElement(String locatorId, String locatorString) {
        WebElement webElement = null;
        WebDriverWait wait = new WebDriverWait(webDriver, 10);
        switch (locatorId) {
        case "id":
            webElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(locatorString)));
            break;
        case "name":
            webElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.name(locatorString)));
            break;
        case "className":
            webElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className(locatorString)));
            break;
        case "cssSelector":
            webElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(locatorString)));
            break;
        case "linkText":
            webElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.linkText(locatorString)));
            break;
        case "partialLinkText":
            webElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.partialLinkText(locatorString)));
            break;
        case "tagName":
            webElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.tagName(locatorString)));
            break;
        case "xpath":
            webElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(locatorString)));
            break;
        default:
            break;
        }
        return webElement;
    }

    public void clickElement(WebElement webElement) {
        webElement.click();
    }

    public void enterInput(WebElement webElement, String inputValue) {
        webElement.sendKeys(inputValue);
    }

    public void dragAndDrop(String locatorId, String locatorString, String inputValue) {
        WebElement webElement1 = findElement(locatorId, locatorString);
        WebElement webElement2 = findElement(locatorId, inputValue);
        (new Actions(webDriver)).dragAndDrop(webElement1, webElement2).perform();
    }

    public static void screenShot(String inputValue) {
        File scrFile = ((TakesScreenshot) webDriver).getScreenshotAs(OutputType.FILE);
        try {
            FileUtils.copyFile(scrFile, new File("/SelScripts/TestAuto_Framework/screenshots/" + inputValue + ".png"));
        } catch (IOException e) {
            logger.error("File not found!");
        }
    }

    @Test(dataProvider = "data")
    public void excuseKeyword(String keyword, String locatorId, String locatorString, String inputValue,
            String numbersOfInputParam, String numbersOfOutputPram, String inputQuery, String expexted) {
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

        String[][] keywords = new String[totalRows][MAX];
        for (int i = 0; i < totalRows; i++) {
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
