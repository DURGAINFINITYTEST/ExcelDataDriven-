import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingXL {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		// workbook---sheet--rows--cells

		XSSFWorkbook workbook = new XSSFWorkbook();

		XSSFSheet sheet = workbook.createSheet("emp info");

		// data
		Object empdata[][] = { { "EMPID", "NAME", "JOB" }, { 101, "prasad", "eng" }, { 102, "mani", "tl" },
				{ 103, "durga", "Manager" }, };
		// using for loop for count the rows and columns

		int rows = empdata.length;
		int column = empdata[0].length;

		System.out.println(rows);
		System.out.println(column);

		for (int i = 0; i < rows; i++) {

			XSSFRow row = sheet.createRow(i);

			for (int c = 0; c < column; c++) {
				XSSFCell cell = row.createCell(c);
				Object value = empdata[i][c];
				if (value instanceof String)
					cell.setCellValue((String) value);
				if (value instanceof String)
					cell.setCellValue((Integer) value);
				if (value instanceof String)
					cell.setCellValue((Boolean) value);
			}

		}

	}

}
