package fsoft.com.vn.automationtestframework;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// TODO: Auto-generated Javadoc
/**
 * The Class ReadScript.
 */
public class ReadScript {
	
	/**
	 * Read excel data.
	 *
	 * @param file the file
	 * @param numColumn the num column
	 * @return the object[][]
	 */
	@SuppressWarnings("deprecation")
	public static Object[][] readExcelData(File file, int numColumn) {

		Object[][] arrayExcelData = null;
		List<String> listInputParam;
		List<String> listOutputParam;
		try {
			FileInputStream fileInputStream = new FileInputStream(file);

			// Create Workbook instance holding reference to .xlsx file
			@SuppressWarnings("resource")
			XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheet("TC01_UI");

			arrayExcelData = new Object[sheet.getLastRowNum()][numColumn];
			String titleScript;
			Cell cell;
			int numInputParam;
			int numOutputParam;
			String title;
			for (int i = 0; i < arrayExcelData.length; i++) {

				for (int j = 0; j < sheet.getRow(0).getLastCellNum(); j++) {
					titleScript = String.valueOf(sheet.getRow(0).getCell(j)).trim();
					cell = sheet.getRow(i + 1).getCell(j);
					cell.setCellType(Cell.CELL_TYPE_STRING);
					switch (titleScript) {
					case "Keyword":
						arrayExcelData[i][0] = cell.getStringCellValue();
						break;
					case "Locator ID":
						arrayExcelData[i][1] = cell.getStringCellValue();
						break;
					case "Locator String":
						arrayExcelData[i][2] = cell.getStringCellValue();
						break;
					case "Input Value":
						arrayExcelData[i][3] = cell.getStringCellValue();

						break;
					case "Row #":
						arrayExcelData[i][4] = cell.getStringCellValue();
						break;
					case "Column #":
						arrayExcelData[i][5] = cell.getStringCellValue();
						break;
					case "#Input Param":
						if (!"".equals(cell.getStringCellValue())) {
							listInputParam = new ArrayList<String>();
							numInputParam = Integer.parseInt(cell.getStringCellValue());
							for (int j2 = 1; j2 <= numInputParam; j2++) {
								title = String.format("Input Param %s", j2);
								for (int j3 = 0; j3 < sheet.getRow(0).getLastCellNum(); j3++) {
									titleScript = String.valueOf(sheet.getRow(0).getCell(j3)).trim();
									cell = sheet.getRow(i + 1).getCell(j3);
									cell.setCellType(Cell.CELL_TYPE_STRING);
									if (title.equals(titleScript)) {
										listInputParam.add(cell.getStringCellValue());
										break;
									}
								}
							}
							arrayExcelData[i][6] = listInputParam;
						} else {
							arrayExcelData[i][6] = new ArrayList<String>();
						}
						break;
					case "#Output Param":
						if (!"".equals(cell.getStringCellValue())) {
							listOutputParam = new ArrayList<String>();
							numOutputParam = Integer.parseInt(cell.getStringCellValue());

							for (int j2 = 1; j2 <= numOutputParam; j2++) {
								title = String.format("Output Param %s", j2);
								for (int j3 = 0; j3 < sheet.getRow(0).getLastCellNum(); j3++) {
									titleScript = String.valueOf(sheet.getRow(0).getCell(j3)).trim();
									cell = sheet.getRow(i + 1).getCell(j3);
									cell.setCellType(Cell.CELL_TYPE_STRING);
									if (title.equals(titleScript)) {
										listOutputParam.add(cell.getStringCellValue());
										break;
									}
								}
							}
							arrayExcelData[i][7] = listOutputParam;
						} else {
							arrayExcelData[i][7] = new ArrayList<String>();
						}
						break;
					default:
						break;
					}
				}

			}

			fileInputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return arrayExcelData;
	}

	/**
	 * Store data return.
	 *
	 * @param file the file
	 * @param rowNumber the row number
	 * @param outputParam the output param
	 * @throws FileNotFoundException the file not found exception
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public static void store_data_return(File file, int rowNumber, List<String> outputParam)
			throws FileNotFoundException, IOException {
		FileInputStream fileInputStream = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet sheet = workbook.getSheet("TC01_UI");
		int colNumber = sheet.getRow(0).getLastCellNum();
		String titleScript;
		String title;
		int lengthOutput = outputParam.size();
		Cell cell;
		int indexOP = 1;
		for (int i = 0; i < colNumber; i++) {
			titleScript = String.valueOf(sheet.getRow(0).getCell(i)).trim();
			if ("#Output Param".equals(titleScript)) {
				cell = sheet.getRow(rowNumber).getCell(i);
				cell.setCellValue(String.valueOf(lengthOutput));
				break;
			}
		}
		for (String item : outputParam) {

			title = String.format("Output Param %s", indexOP);
			for (int i = 0; i < colNumber; i++) {
				titleScript = String.valueOf(sheet.getRow(0).getCell(i)).trim();
				if (title.equals(titleScript)) {
					cell = sheet.getRow(rowNumber).getCell(i);
					cell.setCellValue(item);
					break;
				}
			}
			indexOP++;
		}
		try {
			FileOutputStream outputStream = new FileOutputStream(file);
			workbook.write(outputStream);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
}
