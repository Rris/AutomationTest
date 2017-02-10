package fsoft.com.vn.automationtestframework;

import java.io.File;
import java.io.IOException;
import java.rmi.*;
import java.rmi.server.*;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;


// TODO: Auto-generated Javadoc
/**
 * The Class SeleniumImpl.
 */
public class SeleniumImpl implements Selenium{
	
	/** The driver. */
	WebDriver driver;
	
	/** The wait. */
	WebDriverWait wait;
	
	/** The element. */
	WebElement element;
	
	/** The by. */
	By by;
	
	/** The flag. */
	static boolean flag = false;
    
    /** The data index. */
    static int dataIndex = 0;
    
    /** The from line. */
    static int fromLine = 0;
    
    /** The loop times. */
    static int loopTimes = 0;
    
    /** The sub data. */
    static List<Map<String, String>> subData = null; 
	
	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#getBy(java.lang.String, java.lang.String)
	 */
	@Override
	public By getBy(String locatorID, String locatorString) {
		
		if ("id".equals(locatorID)) {
			by = By.id(locatorString);
		} else if ("linkText".equals(locatorID)) {
			by = By.linkText(locatorString);
		} else if ("name".equals(locatorID)) {
			by = By.name(locatorString);
		} else if ("xpath".equals(locatorID)) {
			by = By.xpath(locatorString);
		}else if("cssSelector".equals(locatorID)){
			by = By.cssSelector(locatorString);
		}
		return by;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#setup_browser(java.lang.String)
	 */
	@Override
	public boolean setup_browser(String inputValue) {
		try {
			if ("Chrome".equals(inputValue)) {
				System.setProperty("webdriver.chrome.driver", "D:\\Software\\Selenium\\chromedriver.exe");
				driver = new ChromeDriver();
			} else if ("Firefox".equals(inputValue)) {
				System.setProperty("webdriver.firefox.marionette",
						"D:\\Software\\Selenium\\FirefoxDriver\\geckodriver.exe");
				driver = new FirefoxDriver();
			}
		} catch (Exception e) {
			return false;
		}
		
		return true;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#explicit_wait(java.lang.String)
	 */
	@Override
	public boolean explicit_wait(String inputValue) {
		wait = new WebDriverWait(driver, Integer.parseInt(inputValue));
		return true;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#load_page(java.lang.String)
	 */
	@Override
	public boolean load_page(String url) {
		try {
			driver.get(url);
		} catch (Exception e) {
			return false;
		}
		return true;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#click_element(java.lang.String, java.lang.String)
	 */
	@Override
	public boolean click_element(String locatorID, String locatorString) {
		try {
			element = wait.until(ExpectedConditions.visibilityOfElementLocated(getBy(locatorID, locatorString)));
			element.click();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#enter_input(java.lang.String, java.lang.String, java.lang.String)
	 */
	@Override
	public boolean enter_input(String locatorID, String locatorString, String inputValue) {
		try {
			element = wait.until(ExpectedConditions.visibilityOfElementLocated(getBy(locatorID, locatorString)));
			element.clear();
			element.sendKeys(inputValue);
		} catch (Exception e) {
			return false;
		}
		return true;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#select_drop_down(java.lang.String, java.lang.String, java.lang.String)
	 */
	@Override
	public boolean select_drop_down(String locatorID, String locatorString, String inputValue) {
		try {
			element = wait.until(ExpectedConditions.visibilityOfElementLocated(getBy(locatorID, locatorString)));
			Select select = new Select(element);
			select.selectByVisibleText(String.valueOf(inputValue));
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#verify_element_present(java.lang.String, java.lang.String, java.lang.String)
	 */
	@Override
	public boolean verify_element_present(String locatorID, String locatorString, String inputValue)
			 {
		try {
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.linkText("Đăng nhập")));
		} catch (NoSuchElementException e2) {
			if (Boolean.parseBoolean(inputValue)) {
				return false;
			} else if (!Boolean.parseBoolean(inputValue)) {
				return true;
			}
		} catch (TimeoutException e3) {
			if (Boolean.parseBoolean(inputValue)) {
				return false;
			} else if (!Boolean.parseBoolean(inputValue)) {
				return true;
			}
		}
		if (!Boolean.parseBoolean(inputValue)) {
			return false;
		}
		return true;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#enter_input_hidden(java.lang.String, java.lang.String, java.lang.String)
	 */
	@Override
	public boolean enter_input_hidden(String locatorID, String locatorString, String inputValue)
		{
		try {
			JavascriptExecutor jse = (JavascriptExecutor) driver;
			((JavascriptExecutor) jse).executeScript("arguments[0].type ='text';",
					wait.until(ExpectedConditions.presenceOfElementLocated(getBy(locatorID, locatorString))));
			wait.until(ExpectedConditions.presenceOfElementLocated(getBy(locatorID, locatorString))).clear();
			wait.until(ExpectedConditions.presenceOfElementLocated(getBy(locatorID, locatorString)))
					.sendKeys(inputValue);
		} catch (NoSuchElementException e) {
			return false;
		} catch (TimeoutException e) {
			return false;
		}
		return true;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#count_drop_down(java.lang.String, java.lang.String, java.lang.String)
	 */
	@Override
	public boolean count_drop_down(String locatorID, String locatorString, String inputValue) {
		Select se;
		try {
			se = new Select(wait.until(ExpectedConditions.visibilityOfElementLocated(getBy(locatorID, locatorString))));

		} catch (NoSuchElementException e) {
			return false;
		}
		List<WebElement> l = se.getOptions();
		l.size();
		return true;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#verify_text(java.lang.String, java.lang.String, java.lang.String)
	 */
	@Override
	public boolean verify_text(String locatorID, String locatorString, String inputValue) {
		String result;
		try {
			result = wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("customername"))).getText();
			if (inputValue.equals(result)) {
				return true;
			}

		} catch (NoSuchElementException e) {
			return false;
		}
		return false;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#verify_text_contains(java.lang.String, java.lang.String, java.lang.String)
	 */
	@Override
	public boolean verify_text_contains(String locatorID, String locatorString, String inputValue)
			 {
		String result;
		try {
			result = wait.until(ExpectedConditions.visibilityOfElementLocated(getBy(locatorID, locatorString)))
					.getText();
			if (inputValue.equals(result)) {
				return true;
			}

		} catch (NoSuchElementException e) {
			return false;
		}
		return false;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#verify_field_text(java.lang.String, java.lang.String, java.lang.String)
	 */
	@Override
	public boolean verify_field_text(String locatorID, String locatorString, String inputValue) {
		String result;
		try {
			result = wait.until(ExpectedConditions.visibilityOfElementLocated(getBy(locatorID, locatorString)))
					.getText();
			if (inputValue.equals(result)) {
				return true;
			}

		} catch (NoSuchElementException e) {
			return false;
		}
		return false;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#verify_table_cell_text(java.lang.String, java.lang.String, java.lang.String, java.lang.String, java.lang.String)
	 */
	@Override
	public boolean verify_table_cell_text(String locatorID, String locatorString, String inputValue, String row,
			String col) {
		String locatorS = locatorString + "/tbody/tr[" + row + "]/td[" + col + "]";
		String result;
		try {
			result = wait.until(ExpectedConditions.visibilityOfElementLocated(getBy(locatorID, locatorS))).getText();
			if (inputValue.equals(result)) {
				return true;
			}

		} catch (NoSuchElementException e) {
			return false;
		}
		return false;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#table_row_sel(java.lang.String, java.lang.String, java.lang.String, java.lang.String)
	 */
	@Override
	public int table_row_sel(String locatorID, String locatorString, String inputParam, String col)
			{
		int rowNo = 0;
		String locatorS;
		try {
			int rows = driver.findElements(getBy(locatorID, locatorString + "/tbody/tr")).size();
			System.out.println("Size: " + rows);
			for (int i = 1; i < rows; i++) {
				locatorS = locatorString + "/tbody/tr[" + (i + 1) + "]/td[" + col + "]";
				WebElement row = wait.until(ExpectedConditions.visibilityOfElementLocated(getBy(locatorID, locatorS)));
				if (row.getText().trim().contains(inputParam)) {
					rowNo = i;
					break;
				}
			}
		} catch (NoSuchElementException e) {
			return 0;
		}
		return rowNo;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#store_cell_data(java.lang.String, java.lang.String, java.lang.String, java.lang.String)
	 */
	@Override
	public String store_cell_data(String locatorID, String locatorString, String col, String row)
			{
		
		String locatorS = locatorString + "/tbody/tr[" + (Integer.parseInt(row) + 1) + "]/td[" + (Integer.parseInt(col) + 1) + "]";
		String result;
		try {
			result = wait.until(ExpectedConditions.visibilityOfElementLocated(getBy(locatorID, locatorS))).getText();
		} catch (NoSuchElementException e) {
			return null;
		}
		return result;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#switch_alert_ok()
	 */
	@Override
	public boolean switch_alert_ok() {
		try {
			Alert alert = wait.until(ExpectedConditions.alertIsPresent());
			System.out.println(alert.getText());
			alert.accept();
		} catch (Exception e) {
			return false;
		}
		return true;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#compare_JSAlert_msg(java.lang.String)
	 */
	@Override
	public boolean compare_JSAlert_msg(String inputValue) {
		try {
			Alert alert = wait.until(ExpectedConditions.alertIsPresent());
			if(inputValue.equals(alert.getText())){
				alert.accept();
				return true;
			}else{
				alert.accept();
				return false;
			}
		} catch (Exception e) {
			return false;
		}
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#page_scroll(java.lang.String)
	 */
	@Override
	public boolean page_scroll(String inputValue) {
		try {
			JavascriptExecutor jse = (JavascriptExecutor)driver;
			jse.executeScript("window.scrollBy(0, " + inputValue + " )", "");
			
		} catch (Exception e) {
			return false;
		}
		return true;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#screen_shot(java.lang.String)
	 */
	@Override
	public boolean screen_shot(String inputValue) {
		
		try {
			File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
			FileUtils.copyFile(scrFile, new File("D:\\Selenium\\Data\\Screen\\" + inputValue + ".png"));
		} catch (IOException e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#wait_time(java.lang.String)
	 */
	@Override
	public void wait_time(String inputValue) throws NumberFormatException, InterruptedException {
		Thread.sleep(Integer.parseInt(inputValue));
	}

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#store_data(java.lang.String, java.lang.String)
	 */
	@Override
	public String store_data(String locatorID, String locatorString) {
		String result;
		try {
			result = (String) wait.until(ExpectedConditions.visibilityOfElementLocated(getBy(locatorID, locatorString))).getText();
		
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
		return result;
	}
	
	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#switchFrame(java.lang.String)
	 */
	public void switchFrame(String inputValue) {
        switch (inputValue) {
        case "default":
            driver.switchTo().defaultContent();
            break;
        case "contentFrame":
        	driver.switchTo().parentFrame();
        default:
            break;
        }
    }
	
	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#deselectDropDown(java.lang.String, java.lang.String)
	 */
	public void deselectDropDown(String locatorId, String locatorString) {
        Select select = new Select((WebElement) ExpectedConditions.visibilityOfElementLocated(getBy(locatorId, locatorString)));
        select.deselectAll();
        // select.deselectByIndex(0);
    }

	/* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#closeDriver()
	 */
    public void closeDriver() {
    	driver.close();
    }

    /* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#quitDriver()
	 */
    public void quitDriver() {
    	driver.quit();
    }

    /* (non-Javadoc)
	 * @see fsoft.com.vn.AutomationTestWithLookup.Selenium#sendKey()
	 */
    public void sendKey() {
        Actions builder = new Actions(driver);
        builder.keyDown(Keys.TAB).perform();
    }
}
