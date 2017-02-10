package fsoft.com.vn.automationtestframework;

import org.testng.annotations.Test;

import org.testng.annotations.DataProvider;
import org.testng.annotations.BeforeClass;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.rmi.RemoteException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.testng.Assert;
import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;

// TODO: Auto-generated Javadoc
/**
 * The Class TestNGClass.
 */
public class TestNGClass {
	
	/** The se impl. */
	SeleniumImpl seImpl;
	
	/** The url. */
	String url;
	
	/** The i. */
	int i = 1;

	/**
	 * Verify automation test.
	 *
	 * @param keyWord the key word
	 * @param locatorID the locator ID
	 * @param locatorString the locator string
	 * @param inputValue the input value
	 * @param row the row
	 * @param col the col
	 * @param listInput the list input
	 * @param listOutput the list output
	 * @throws FileNotFoundException the file not found exception
	 * @throws IOException Signals that an I/O exception has occurred.
	 * @throws NumberFormatException the number format exception
	 * @throws InterruptedException the interrupted exception
	 */
	@Test(dataProvider = "dp")
	public void VerifyAutomationTest(String keyWord, String locatorID, String locatorString, String inputValue,
			String row, String col, List<String> listInput, List<String> listOutput)
			throws FileNotFoundException, IOException, NumberFormatException, InterruptedException {

		switch (keyWord) {
		case "setup_browser":
			seImpl.setup_browser(inputValue);
			break;
		case "explicit_wait":
			Assert.assertTrue(seImpl.explicit_wait(inputValue));
			break;
		case "load_page":
			Assert.assertTrue(seImpl.load_page(inputValue));
			break;
		case "click_element":
			Assert.assertTrue(seImpl.click_element(locatorID, locatorString));
			break;
		case "enter_input":
			Assert.assertTrue(seImpl.enter_input(locatorID, locatorString, inputValue));
			break;
		case "select_drop_down":
			Assert.assertTrue(seImpl.select_drop_down(locatorID, locatorString, inputValue));
			break;
		case "verify_element_present":
			Assert.assertTrue(seImpl.verify_element_present(locatorID, locatorString, inputValue));
			break;
		case "enter_input_hidden":
			Assert.assertTrue(seImpl.enter_input_hidden(locatorID, locatorString, inputValue));
			break;
		case "count_drop_down":
			Assert.assertTrue(seImpl.count_drop_down(locatorID, locatorString, inputValue));
			break;
		case "verify_text":
			Assert.assertTrue(seImpl.verify_text(locatorID, locatorString, inputValue));
			break;
		case "verify_field_text":
			Assert.assertTrue(seImpl.verify_field_text(locatorID, locatorString, inputValue));
			break;
		case "verify_text_contains":
			Assert.assertTrue(seImpl.verify_text_contains(locatorID, locatorString, inputValue));
			break;

		case "verify_table_cell_text":
			Assert.assertTrue(seImpl.verify_table_cell_text(locatorID, locatorString, inputValue, row, col));
			break;

		case "table_row_sel":
			List<String> l_table_row_sel = new ArrayList<String>();
			l_table_row_sel.add(String.valueOf(seImpl.table_row_sel(locatorID, locatorString, listInput.get(0), col)));
			ReadScript.store_data_return(new File(url), i, l_table_row_sel);
			break;
		case "store_cell_data":
			List<String> l_store_cell_data = new ArrayList<String>();
			l_store_cell_data
					.add(String.valueOf(seImpl.store_cell_data(locatorID, locatorString, col, listInput.get(0))));
			ReadScript.store_data_return(new File(url), i, l_store_cell_data);
			break;
		case "switch_alert_ok":
			Assert.assertTrue(seImpl.switch_alert_ok());
			break;
		case "compare_JSAlert_msg":
			Assert.assertTrue(seImpl.compare_JSAlert_msg(inputValue));
			break;
		case "page_scroll":
			Assert.assertTrue(seImpl.page_scroll(inputValue));
			break;
		case "screen_shot":
			Assert.assertTrue(seImpl.screen_shot(inputValue));
			break;
		case "wait_time":
			seImpl.wait_time(inputValue);
			break;
		case "store_data":
			String result = seImpl.store_data(locatorID, locatorString);
			String result1 = seImpl.store_data("linkText", "Đăng ký");
			List<String> l_store_data = new ArrayList<String>();
			l_store_data.add(result);
			l_store_data.add(result1);
			ReadScript.store_data_return(new File(url), i, l_store_data);
			break;
		
		default:
			break;
		}
		i++;
	}

	/**
	 * Dp.
	 *
	 * @return the object[][]
	 */
	@DataProvider
	public Object[][] dp() {
		Object[][] arrayList = ReadScript.readExcelData(new File(url), 8);
		return arrayList;
	}

	/**
	 * Before class.
	 *
	 * @throws RemoteException the remote exception
	 */
	@BeforeClass
	public void beforeClass() throws RemoteException {
		seImpl = new SeleniumImpl();
		url = "D:\\Selenium\\Data\\test-data-1.xlsx";
	}

	/**
	 * After class.
	 *
	 * @param result the result
	 */
	@AfterMethod
	public void afterClass(ITestResult result) {
		Date date = new Date();
		SimpleDateFormat simpleDateFormat = new SimpleDateFormat("EE MM-DD-YYYY hh:mm:ss");
		String status = null;
		String msg = null;
		msg = String.format("TestNG Debugger : %s() running with parameters %s.",
				result.getMethod().getConstructorOrMethod().getName(), Arrays.toString(result.getParameters()));
		try {
			if (result.getStatus() == ITestResult.SUCCESS) {
				status = "PASSED";
			} else if (result.getStatus() == ITestResult.FAILURE) {
				status = "FAILED";
			} else if (result.getStatus() == ITestResult.SKIP) {
				status = "SKIPED";
			}
			SaveResult.saveResult("D:/result.txt", simpleDateFormat.format(date), msg, status);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
