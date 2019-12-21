import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 


class nullcheck {

	public static void main(String args[]) {
		HSSFWorkbook workBook = null;
		try {
			FileInputStream fis = new FileInputStream("./calctest.xls");
			workBook = new HSSFWorkbook(fis);
			fis.close();
		}
		catch(Exception e) {
			System.out.println(e.toString());
		}
		HSSFSheet sheet = workBook.getSheetAt(0);
/*
		HSSFRow row = sheet.getRow(0);
		row.getCell(0).setCellValue(1);
		row.getCell(1).setCellValue(3);
		String fum = row.getCell(2).getCellFormula();
		System.out.println(fum);
		row.getCell(2).setCellFormula(fum);
*/
		HSSFRow row = sheet.getRow(1);
		if (row == null) {
			System.out.println("rowÇÕÉkÉãÇ≈Ç∑");
		}
		else {
			HSSFCell cell = row.getCell(0);
			if (cell == null) {
				System.out.println("cellÇÕÉkÉãÇ≈Ç∑");
			}
			else {
				String s = cell.getStringCellValue();
				if (s==null) {
					System.out.println("cellÇÃíÜêgÇÕÉkÉãÇ≈Ç∑");
				}
			}
		}
/*	
		try {
			FileOutputStream fos = new FileOutputStream("./calctest.xls");
			workBook.write(fos);
			fos.close();
			System.out.println("Done");
		}
		catch(Exception e) {
			System.out.println(e.toString());
		}
*/
	}
}
