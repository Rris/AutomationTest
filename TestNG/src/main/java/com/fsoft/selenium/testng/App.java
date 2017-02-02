package com.fsoft.selenium.testng;

import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {

    private static int keyword;
    private static int locatorId;
    private static int locatorString;
    private static int inputValue;
    private static int numbersOfInputParam;
    private static int numbersOfOutputParam;
    private static int inputQuery;
    private static int expected;

    public static String[][] getDataFromExcelFile(String filePath, String sheetName) {
        String[][] arrKeywords = null;
        try {
            FileInputStream fileInput = new FileInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(fileInput);
            XSSFSheet sheet = workbook.getSheet(sheetName);
            int totalRows = sheet.getPhysicalNumberOfRows();
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
                        case "Locator String":
                            locatorString = i;
                            break;
                        case "Input Value":
                            inputValue = i;
                            break;
                        case "#Input Param":
                            numbersOfInputParam = i;
                            break;
                        case "#Output Param":
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
                        arrKeywords[row.getRowNum() - 1][i] = row.getCell(i).getStringCellValue();
                    }
                }
            }
            workbook.close();
        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
        return arrKeywords;
    }

    public static void main(String[] args) {
        getDataFromExcelFile("D:/4.1.RITE20_ModuleName_UI_FunctionName_TestCases.xlsx", "TC01_UI");
    }
}
