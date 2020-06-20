package RunnerFile;

import prac_ro.ReadGuru99ExcelFile;
import prac_ro.exclData;

public class TestRunner {

	public static void main(String args[]) throws Exception
	{
		exclData.readexcl();
		System.out.println(" --------------------------------------------     ------------------------------------------ ");
		ReadGuru99ExcelFile.readexcel();
		
	}
}
