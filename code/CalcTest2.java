import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 


class CalcTest {

	public static void main(String args[]) {
		String mode = args[0];
		Workbook workBook = null;
		try {
			if (mode.equals("2003")) {
				FileInputStream fis = new FileInputStream("./calctest.xls");
				workBook = new HSSFWorkbook(fis);
				fis.close();
			}
			else {
				FileInputStream fis = new FileInputStream("./calctest.xlsx");
				workBook = new XSSFWorkbook(fis);
				fis.close();
			}
		}
		catch(Exception e) {
			System.out.println(e.toString());
		}
		Sheet sheet = workBook.getSheetAt(0);
		Row row = sheet.getRow(0);
		row.getCell(0).setCellValue(1);
		row.getCell(1).setCellValue(7);
		String fum = row.getCell(2).getCellFormula();
		System.out.println(fum);
		row.getCell(2).setCellFormula(fum);
		try {
			
			FileOutputStream fos = null;
			if (mode.equals("2003")) {
				fos = new FileOutputStream("./calctest.xls");
			}
			else {
				fos = new FileOutputStream("./calctest.xlsx");
			}
			workBook.write(fos);
			fos.close();
			System.out.println("Done");
		}
		catch(Exception e) {
			System.out.println(e.toString());
		}
	}
}
