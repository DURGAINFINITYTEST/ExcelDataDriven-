import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DatadrivenIterator {

	public static void main(String[] args) throws Exception {

		String path = "C:\\Users\\katak\\eclipse-workspace\\PavanDatadriven\\pavandata.xlsx";

		FileInputStream fis = new FileInputStream(path);

		XSSFWorkbook xss = new XSSFWorkbook(fis);
		XSSFSheet sheet = xss.getSheetAt(0);
		Iterator<Row> iterator = sheet.rowIterator();
		while (iterator.hasNext()) {
			XSSFRow row = (XSSFRow) iterator.next();
			Iterator cellIterator = row.cellIterator();
			if (cellIterator.hasNext()) {

				XSSFCell cell = (XSSFCell) cellIterator.next();

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

			}
		}

	}

}
