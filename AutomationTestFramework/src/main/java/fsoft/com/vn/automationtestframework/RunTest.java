package fsoft.com.vn.automationtestframework;

import org.testng.TestListenerAdapter;
import org.testng.TestNG;

// TODO: Auto-generated Javadoc
/**
 * The Class RunTest.
 */
public class RunTest {
	public static void main(String[] args) {
		TestListenerAdapter tla = new TestListenerAdapter();
		TestNG testng = new TestNG();
		testng.setTestClasses(new Class[] { TestNGClass.class });
		testng.addListener(tla);
		testng.run();
	}

}
