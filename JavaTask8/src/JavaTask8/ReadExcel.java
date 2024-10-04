package JavaTask8;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	public static void main(String[] args) {
		try {
			getRowCount();
		} catch (IOException e) {

			e.printStackTrace();
		}
	}

	public static void getRowCount() throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook(".\\excel\\Employeedata.xlsx");
		XSSFSheet sheet = workbook.getSheet("Sheet 1");
		int rowcount = sheet.getPhysicalNumberOfRows();
		System.out.println(rowcount);

		sheet = workbook.getSheet("Sheet 1");
		String celldata = sheet.getRow(0).getCell(2).getStringCellValue();
		System.out.println(celldata);

		sheet = workbook.getSheet("Sheet 1");
		String celldata1 = sheet.getRow(3).getCell(1).getStringCellValue();
		System.out.println(celldata1);

	}

}
