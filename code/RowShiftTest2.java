import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * Row�ړ��̃e�X�g
 */
public class RowShiftTest2 {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    // ���[�N�u�b�N��ǂݍ���
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./ShiftRow.xls" : "./ShiftRow.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + e.toString());
      return;
    }
    // Sheet1�̓ǂݍ��� 
    Sheet sheet = workBook.getSheetAt(0);
    // 4�s�ڂ��폜
    sheet.removeRow(sheet.getRow(3));
    // Row�̈ړ�
    sheet.shiftRows(4, 9, -1, true, true);
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./Book1.xls" : 
                      "./Book1.xlsx");
      workBook.write(out);
    }catch(IOException e){
      System.out.println("�u�b�N�̏������݂Ɏ��s���܂����B\n" + e.toString());
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println("�u�b�N�̏������݂Ɏ��s���܂����B\n" + e.toString());
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
    new RowShiftTest2().Run(args[0]);
  }
}
