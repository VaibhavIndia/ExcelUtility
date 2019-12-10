package com.main;

import java.io.File;
import java.util.ArrayList;

import com.excelutility.VaibhavExelUtility;

public class MainFile {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		File fl = new File (System.getProperty("user.dir")+"\\src\\com\\exceldata\\data1.xlsx");
		ArrayList<String> a  = VaibhavExelUtility.getdata(fl, "Sheet1", "Testcase", "case2");
		for (String s :a)
		{
			System.out.println(s);
		}
		System.out.println("done");
		
		
		String sample = VaibhavExelUtility.getDataFromRequiredRowAndCell(fl, "Sheet1", 2 ,3);
		System.out.println(sample);
		System.out.println("---------");
		String b = VaibhavExelUtility.getCell(fl,"Sheet1", "G2");
		System.out.println(b);
		
		
		
	}

}
