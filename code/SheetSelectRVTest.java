import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * �V�[�g�I���̃e�X�g
 */
public class SheetSelectRVTest {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
		Workbook workBook = null; 
   // ���[�N�u�b�N��ǂݍ���
    try {
      FileInputStream fis = new FileInputStream( mode.equals("2003") ? "./book1.xls" : "./book1.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + e.toString());
      return;
    }
		Sheet sheet = workBook.getSheetAt(0);
		Row row = sheet.getRow(0);
		if (row == null) {
			System.out.println("Row[0]�͑��݂��Ȃ�");
			Row rown = sheet.createRow(0);
			Cell cell = rown.createCell(0);
			System.out.println("����" + cell.getRowIndex() + "�s�ڂ�" + cell.getColumnIndex() + "�Ԗڂ̃Z���ł�");
			cell.setAsActiveCell();
		}
		else {
			System.out.println("Row[0]�͑��݂���");
			Cell cell = row.getCell(0,Row.CREATE_NULL_AS_BLANK);
			System.out.println("����" + cell.getRowIndex() + "�s�ڂ�" + cell.getColumnIndex() + "�Ԗڂ̃Z���ł�");
		}
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./Book1.xls" : 
                      "./Book1.xlsx");
      workBook.write(out);
    }catch(IOException e){
      System.out.println(e.toString());
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println(e.toString());
      }
    }
    System.out.println("done!");
  }
  /** �G���g���[�|�C���g */

  public static void main(String[] args) {

    if (args.length != 1) {
      System.out.println("�G���[�F���[�h���w�肵�ĉ������B");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }
    new SheetSelectRVTest().Run(args[0]);
  }
}
