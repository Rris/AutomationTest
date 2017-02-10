package fsoft.com.vn.automationtestframework;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.PrintWriter;

// TODO: Auto-generated Javadoc
/**
 * The Class SaveResult.
 */
public class SaveResult {
	
	/**
	 * Save result.
	 *
	 * @param url the url
	 * @param time the time
	 * @param nameTestMethod the name test method
	 * @param status the status
	 */
	public static void saveResult(String url, String time, String nameTestMethod, String status) {
		FileOutputStream fos;
		try {
			fos = new FileOutputStream(url, true);
			PrintWriter pw = new PrintWriter(fos);
			pw.println( time);
			pw.println( nameTestMethod);
			pw.println( "Status: " + status);
			pw.println("\n*****************************************************************************");
			pw.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}

	}
}
