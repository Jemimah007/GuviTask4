package JavaTask8;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public static void main(String[] args) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Sheet 1");
		Object data[][] = { { "Name", "Age", "Email" }, { "John Doe", "30", "john@test.com" },
				{ "Jane Doe", "28", "john@test.com" }, { "Bob Smith", "35", "jachy@example.com" },
				{ "Swapnil", "37", "swapnil@example.com" } };

		int rows = data.length;
		int cols = data[0].length;

		for (int r = 0; r < rows; r++) {
			XSSFRow row = sheet.createRow(r);
			for (int c = 0; c < cols; c++) {
				XSSFCell cell = row.createCell(c);
				Object value = data[r][c];
				if (value instanceof String)
					cell.setCellValue((String) value);
				if (value instanceof Integer)
					cell.setCellValue((Integer) value);
				if (value instanceof Boolean)
					cell.setCellValue((Boolean) value);
			}

		}
		String filepath = ".\\excel\\Employeedata.xlsx";

		FileOutputStream fos = new FileOutputStream(filepath);
		workbook.write(fos);
		fos.close();
		System.out.println("Excel created");

	}

}
