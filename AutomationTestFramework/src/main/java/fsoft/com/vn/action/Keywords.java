package fsoft.com.vn.action;

import java.util.ArrayList;

public abstract class Keywords {
	String locatorId;
	String locatorString;
	String inputValue;
	int row;
	int column;
	ArrayList<Integer> inputParams;
	ArrayList<Integer> outputParams;
	String inputQuery;
	String expected;

	public Keywords() {
		// TODO Auto-generated constructor stub
	}
	
	void execute() {
	}
}
