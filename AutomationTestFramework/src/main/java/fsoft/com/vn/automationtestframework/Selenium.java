package fsoft.com.vn.automationtestframework;

import java.rmi.*;
import java.util.List;
import java.util.Map;

import org.openqa.selenium.*;

// TODO: Auto-generated Javadoc
/**
 * The Interface Selenium.
 */
public interface Selenium{
	
	/**
	 * Gets the by.
	 *
	 * @param locatorID the locator ID
	 * @param locatorString the locator string
	 * @return the by
	 * @throws RemoteException the remote exception
	 */
	public By getBy(String locatorID, String locatorString);
	
	/**
	 * Setup browser.
	 *
	 * @param inputValue the input value
	 * @return true, if successful
	 * @throws RemoteException the remote exception
	 */
	public boolean setup_browser(String inputValue);
	
	/**
	 * Explicit wait.
	 *
	 * @param inputValue the input value
	 * @return true, if successful
	 * @throws RemoteException the remote exception
	 */
	public boolean explicit_wait(String inputValue);
	
	/**
	 * Load page.
	 *
	 * @param url the url
	 * @return true, if successful
	 * @throws RemoteException the remote exception
	 */
	public boolean load_page(String url);
	
	/**
	 * Click element.
	 *
	 * @param locatorID the locator ID
	 * @param locatorString the locator string
	 * @return true, if successful
	 * @throws RemoteException the remote exception
	 */
	public boolean click_element(String locatorID, String locatorString);
	
	/**
	 * Enter input.
	 *
	 * @param locatorID the locator ID
	 * @param locatorString the locator string
	 * @param valueInput the value input
	 * @return true, if successful
	 * @throws RemoteException the remote exception
	 */
	public boolean enter_input(String locatorID, String locatorString, String valueInput);
	
	/**
	 * Select drop down.
	 *
	 * @param locatorID the locator ID
	 * @param locatorString the locator string
	 * @param valueInput the value input
	 * @return true, if successful
	 * @throws RemoteException the remote exception
	 */
	public boolean select_drop_down(String locatorID, String locatorString, String valueInput);
	
	/**
	 * Verify element present.
	 *
	 * @param locatorID the locator ID
	 * @param locatorString the locator string
	 * @param valueInput the value input
	 * @return true, if successful
	 * @throws RemoteException the remote exception
	 */
	public boolean verify_element_present(String locatorID, String locatorString, String valueInput);
	
	/**
	 * Enter input hidden.
	 *
	 * @param locatorID the locator ID
	 * @param locatorString the locator string
	 * @param valueInput the value input
	 * @return true, if successful
	 * @throws RemoteException the remote exception
	 */
	public boolean enter_input_hidden(String locatorID, String locatorString, String valueInput);
	
	/**
	 * Count drop down.
	 *
	 * @param locatorID the locator ID
	 * @param locatorString the locator string
	 * @param valueInput the value input
	 * @return true, if successful
	 * @throws RemoteException the remote exception
	 */
	public boolean count_drop_down(String locatorID, String locatorString, String valueInput);
	
	/**
	 * Verify text.
	 *
	 * @param locatorID the locator ID
	 * @param locatorString the locator string
	 * @param valueInput the value input
	 * @return true, if successful
	 * @throws RemoteException the remote exception
	 */
	public boolean verify_text(String locatorID, String locatorString, String valueInput);
	
	/**
	 * Verify text contains.
	 *
	 * @param locatorID the locator ID
	 * @param locatorString the locator string
	 * @param valueInput the value input
	 * @return true, if successful
	 * @throws RemoteException the remote exception
	 */
	public boolean verify_text_contains(String locatorID, String locatorString, String valueInput);
	
	/**
	 * Verify field text.
	 *
	 * @param locatorID the locator ID
	 * @param locatorString the locator string
	 * @param valueInput the value input
	 * @return true, if successful
	 * @throws RemoteException the remote exception
	 */
	public boolean verify_field_text(String locatorID, String locatorString, String valueInput);
	
	/**
	 * Verify table cell text.
	 *
	 * @param locatorID the locator ID
	 * @param locatorString the locator string
	 * @param valueInput the value input
	 * @param row the row
	 * @param col the col
	 * @return true, if successful
	 * @throws RemoteException the remote exception
	 */
	public boolean verify_table_cell_text(String locatorID, String locatorString, String valueInput, String row,
			String col);
	
	/**
	 * Table row sel.
	 *
	 * @param locatorID the locator ID
	 * @param locatorString the locator string
	 * @param valueInput the value input
	 * @param col the col
	 * @return the int
	 * @throws RemoteException the remote exception
	 */
	public int table_row_sel(String locatorID, String locatorString, String valueInput, String col);
	
	/**
	 * Store cell data.
	 *
	 * @param locatorID the locator ID
	 * @param locatorString the locator string
	 * @param row the row
	 * @param col the col
	 * @return the string
	 * @throws RemoteException the remote exception
	 */
	public String store_cell_data(String locatorID, String locatorString, String row, String col);
	
	/**
	 * Switch alert ok.
	 *
	 * @return true, if successful
	 * @throws RemoteException the remote exception
	 */
	public boolean switch_alert_ok();
	
	/**
	 * Compare JS alert msg.
	 *
	 * @param inputValue the input value
	 * @return true, if successful
	 * @throws RemoteException the remote exception
	 */
	public boolean compare_JSAlert_msg(String inputValue);
	
	/**
	 * Page scroll.
	 *
	 * @param inputValue the input value
	 * @return true, if successful
	 * @throws RemoteException the remote exception
	 */
	public boolean page_scroll(String inputValue);
	
	/**
	 * Screen shot.
	 *
	 * @param keyWord the key word
	 * @throws RemoteException the remote exception
	 */
	public boolean screen_shot(String keyWord);
	
	/**
	 * Wait time.
	 *
	 * @param inputValue the input value
	 * @throws RemoteException the remote exception
	 * @throws NumberFormatException the number format exception
	 * @throws InterruptedException the interrupted exception
	 */
	public void wait_time(String inputValue) throws NumberFormatException, InterruptedException;
	
	/**
	 * Store data.
	 *
	 * @param locatorID the locator ID
	 * @param locatorString the locator string
	 * @return the string
	 * @throws RemoteException the remote exception
	 */
	public String store_data(String locatorID, String locatorString);

	/**
	 * Switch frame.
	 *
	 * @param inputValue the input value
	 */
	public void switchFrame(String inputValue);

	/**
	 * De-select DropDown.
	 *
	 * @param locatorID the locator ID
	 * @param locatorString the locator string
	 */
	public void deselectDropDown(String locatorId, String locatorString);

	/**
	 * Close driver.
	 *
	 */
	public void closeDriver();

	/**
	 * Quit driver.
	 *
	 */
	public void quitDriver();

	/**
	 * Send key.
	 *
	 */
	public void sendKey();
}
