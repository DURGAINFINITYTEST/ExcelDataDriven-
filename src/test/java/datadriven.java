import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class datadriven {//USING FOR LOOP

	public static void main(String[] args) throws Exception {

		String path = "C:\\Users\\katak\\eclipse-workspace\\PavanDatadriven\\pavandata.xlsx";

		FileInputStream fis = new FileInputStream(path);

		XSSFWorkbook xss = new XSSFWorkbook(fis);
		XSSFSheet sheet = xss.getSheetAt(0);

		int rows = sheet.getLastRowNum();
		int column = sheet.getRow(1).getLastCellNum();
		for (int i = 0; i <= rows; i++) {

			XSSFRow row = sheet.getRow(i);
			for (int c = 0; c < column; c++) {

				XSSFCell cell = row.getCell(c);

				switch (cell.getCellType()) {

				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;

				}
				System.out.println("|");

			}
		}

	}
}
