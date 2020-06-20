package prac_ro;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class exclData {

	@SuppressWarnings("incomplete-switch")
	public static void readexcl() throws Exception {

		FileInputStream fis = new FileInputStream(
				"D:\\Projects\\Java Projects\\Project06thApril\\WEbsite\\webauto\\data\\Data.xlsx");

		@SuppressWarnings("resource")
		XSSFWorkbook wb = new XSSFWorkbook(fis);

		XSSFSheet sh = wb.getSheet("Sheet1");

		System.out.println("Hello");

		int rowcount = sh.getPhysicalNumberOfRows();
		// System.out.println("Total number of rows= " + rowcount);

		int colcount = sh.getRow(0).getPhysicalNumberOfCells();
		// System.out.println("Total number of cols " + colcount);

		// System.out.println(sh.getRow(0).getCell(0).getStringCellValue());

		for (int i = 1; i <= rowcount; i++) {

			Row row = sh.getRow(i);
			for (int j = 1; j <= colcount; j++) {

				//System.out.println("Value of i is " + i + " and value of j is " + j);

				Cell cell = row.getCell(j);
				switch (cell.getCellType()) {
				case STRING:
					System.out.print(row.getCell(j).getStringCellValue() + "|| ");
					break;

				case NUMERIC:
					System.out.print((int) row.getCell(j).getNumericCellValue() + "|| ");
					break;

				}

				// System.out.println(sh.getRow(i).getCell(j).getStringCellValue()); //not
				// working

				/*
				 * try {
				 * 
				 * System.out.println("Hey"); //Cell cell = sh.getRow(i).getCell(j);
				 * 
				 * String value; value = sh.getRow(i).getCell(j).getStringCellValue();
				 * System.out.print(value); } catch (Exception e) { System.out.println(e); }
				 */
			}
			//System.out.println("\n \n");
		}

	}
}