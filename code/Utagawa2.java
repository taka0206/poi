import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 

/**
 * �^�킵���͎����Ă݂�B�V���[�Y��2�e
 */
class Utagawa2 {
 
  /** �����̎��s
   * @param mode ���[�h
   */
  public void Run(String mode) {

    FileInputStream fis = null;
    Workbook workBook = null;
    // ���[�N�u�b�N�̓ǂݍ���
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./Book1.xls" : "./Book1.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println(e.toString());
    }
		// �V�[�g���擾
		Sheet sheet1 = null;
		Sheet sheet2 = null;
		Sheet sheet3 = null;

		try {
	    sheet1 = workBook.getSheetAt(0);
		}
		catch( Exception e ) {
      System.out.println("Sheet1�͑��݂��Ȃ��I�I" + e.toString());
			return;
		}
		try {
	    sheet2 = workBook.getSheetAt(1);
		}
		catch( Exception e ) {
      System.out.println("Sheet2�͑��݂��Ȃ��I�I" + e.toString());
			return;
		}
		try {
	    sheet3 = workBook.getSheetAt(2);
		}
		catch( Exception e ) {
      System.out.println("Sheet3�͑��݂��Ȃ��I�I" + e.toString());
			return;
		}
		// �V�[�g�����݂����Row���擾���Ă݂�B
		Row row = sheet1.getRow(0);
		if (row == null) {
      System.out.println("Row�͑��݂��Ȃ��I�I");
			return;
		}
    // Row�����݂����Cell���擾���Ă݂�B
		Cell cell = row.getCell(0);
		if (cell == null) {
      System.out.println("Cell�͑��݂��Ȃ��I�I");
			return;
		}
      System.out.println("�S�Ă͑��݂����B");
  }
  /** �G���g���[�|�C���g */
  public static void main(String[] args) {
    if (args.length != 1) {
      System.out.println("�G���[�F���[�h���w�肵�Ă��������B");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }
    new Utagawa2().Run(args[0]);
  }
}
