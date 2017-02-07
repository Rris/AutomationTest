package com.fsoft.selenium.testng;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
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

    public void enterInput(String locatorId, String locatorString, String inputValue) {
        WebElement webElement = findElement(locatorId, locatorString);
        webElement.sendKeys(inputValue);
    }

    public boolean verifyElementPresent(String inputValue, String locatorId, String locatorString) {
        boolean result = false;
        switch (inputValue.toUpperCase()) {
        case "TRUE":
            if (findElement(locatorId, locatorString) != null) {
                result = true;
            }
            break;
        case "FALSE":
            if (findElement(locatorId, locatorString) == null) {
                result = true;
            }
        default:
            break;
        }
        return result;
    }

    public void enterInputHidden(String locatorId, String locatorString, String inputString) {
        WebElement webElement = findElement(locatorId, locatorString);
        webElement.sendKeys(inputString);
    }

    public void selectDropDown(String locatorId, String locatorString, String inputString) {
        Select select = new Select(findElement(locatorId, locatorString));
        select.selectByVisibleText(inputString);
    }

    public int countDropDown(String locatorId, String locatorString) {
        Select select = new Select(findElement(locatorId, locatorString));
        List<WebElement> lstWebElement = select.getOptions();
        return lstWebElement.size();
    }

    public void clickElement(WebElement webElement) {
        webElement.click();
    }

    public boolean verifyText(String locatorId, String locatorString, String inputString) {
        WebElement webElement = findElement(locatorId, locatorString);
        String content = webElement.getText();
        return content.equals(inputString) ? true : false;
    }

    public boolean verifyTextContains(String locatorId, String locatorString, String inputString) {
        WebElement webElement = findElement(locatorId, locatorString);
        String content = webElement.getText();
        return content.contains(inputString) ? true : false;
    }

    public boolean verifyFieldText(String locatorId, String locatorString, String inputString) {
        WebElement webElement = findElement(locatorId, locatorString);
        String content = webElement.getText();
        return content.equals(inputString) ? true : false;
    }

    public boolean compareTwoVar(String inputParam1, String inputParam2) {
        return inputParam1.equals(inputParam2) ? true : false;
    }

    public boolean verifyTableCellText(String locatorId, String locatorString, String inputValue, int row, int column) {
        WebElement webElement = findElement(locatorId,
                locatorString + "//tr[" + (row + 1) + "]/td[" + (column + 1) + "]");
        String content = webElement.getText();
        return content.equals(inputValue) ? true : false;
    }

    public int tableRowSel(String locatorId, String locatorString, int column, String inputParam) {
        int result = 0;
        List<WebElement> rows = findElements(locatorId, locatorString + "//td[" + (column + 1) + "]");
        for (WebElement row : rows) {
            result++;
            if (row.getText().equals(inputParam)) {
                break;
            }
        }
        return result;
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

    public List<WebElement> findElements(String locatorId, String locatorString) {
        List<WebElement> webElement = null;
        switch (locatorId) {
        case "id":
            webElement = webDriver.findElements(By.id(locatorString));
            break;
        case "name":
            webElement = webDriver.findElements(By.name(locatorString));
            break;
        case "className":
            webElement = webDriver.findElements(By.className(locatorString));
            break;
        case "cssSelector":
            webElement = webDriver.findElements(By.cssSelector(locatorString));
            break;
        case "linkText":
            webElement = webDriver.findElements(By.linkText(locatorString));
            break;
        case "partialLinkText":
            webElement = webDriver.findElements(By.partialLinkText(locatorString));
            break;
        case "tagName":
            webElement = webDriver.findElements(By.tagName(locatorString));
            break;
        case "xpath":
            webElement = webDriver.findElements(By.xpath(locatorString));
            break;
        default:
            break;
        }
        return webElement;
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
        // select.deselectByIndex(0);
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
            enterInput(locatorId, locatorString, inputValue);
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
        String[][] arrData = getDataFromExcelFile("D:/4.1.RITE20_ModuleName_UI_FunctionName_TestCases.xlsx",
                "TC01_UI_TS12");
        Object[][] arrayObject = subDataHandle(arrData);
        return arrayObject;
    }

    @SuppressWarnings("deprecation")
    public String[][] getDataFromExcelFile(String filePath, String sheetName) {
        String[][] arrTestCases = null;
        try {
            FileInputStream fileInput = new FileInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(fileInput);
            XSSFSheet sheet = workbook.getSheet(sheetName);
            int totalRows = sheet.getPhysicalNumberOfRows();
            int totalColumns = sheet.getRow(0).getPhysicalNumberOfCells();
            arrTestCases = new String[totalRows][totalColumns];
            Iterator<Row> rows = sheet.rowIterator();
            while (rows.hasNext()) {
                XSSFRow row = (XSSFRow) rows.next();
                int rowNumber = row.getRowNum();
                for (int i = 0; i < totalColumns; i++) {
                    row.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                    arrTestCases[rowNumber][i] = row.getCell(i).getStringCellValue();
                }
            }
            workbook.close();
        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
        return arrTestCases;
    }

    public String[][] dataHandle(String[][] arrTestCases) {
        int keyword = 0;
        int locatorId = 0;
        int locatorString = 0;
        int inputValue = 0;
        int numbersOfInputParam = 0;
        int numbersOfOutputParam = 0;
        int inputQuery = 0;
        int totalRow = arrTestCases.length;
        for (int i = 0; i < arrTestCases[0].length; i++) {
            switch (arrTestCases[0][i]) {
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
            default:
                break;
            }
        }
        String[][] testCases = new String[totalRow - 1][MAX];
        for (int i = 0; i < totalRow - 1; i++) {
            int pos = 0;
            testCases[i][pos++] = arrTestCases[i + 1][keyword];
            testCases[i][pos++] = arrTestCases[i + 1][locatorId];
            testCases[i][pos++] = arrTestCases[i + 1][locatorString];
            testCases[i][pos++] = arrTestCases[i + 1][inputValue];
            testCases[i][pos] = arrTestCases[i + 1][numbersOfInputParam] == "" ? "0"
                    : arrTestCases[i + 1][numbersOfInputParam];
            int inputParams = Integer.parseInt(testCases[i][pos++]);
            testCases[i][pos] = arrTestCases[i + 1][numbersOfOutputParam] == "" ? "0"
                    : arrTestCases[i][numbersOfOutputParam];
            int outputParams = Integer.parseInt(testCases[i][pos++]);
            testCases[i][pos++] = arrTestCases[i + 1][inputQuery];
            for (int j = 0; j < inputParams; j++) {
                testCases[i][pos++] = arrTestCases[i + 1][numbersOfInputParam + j + 1];
            }
            for (int j = 0; j < outputParams; j++) {
                testCases[i][pos++] = arrTestCases[i + 1][numbersOfOutputParam + j + 1];
            }
        }
        return arrTestCases;
    }

    @SuppressWarnings("unchecked")
    public Map<String, String>[][] subDataHandle(String[][] arrData) {
        int totalRow = arrData.length;
        int totalColumn = arrData[0].length;
        Map<String, String>[][] subData = new Map[totalRow][totalColumn];
        for (int i = 1; i < totalRow; i++) {
            for (int j = 0; j < totalColumn; j++) {
                subData[i - 1][j].put(arrData[0][j], arrData[i][j]);
            }
        }
        return subData;
    }
}
