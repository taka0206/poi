import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * �e�̃��[�N�u�b�N�擾�e�X�g
 */
public class GetParentBookTest {

  /**
   * �V�[�g�ŗL�̏���
   *@param sheet �V�[�g�̎Q��
   */
  public void sheetProc(Sheet sheet) {
    Workbook parentBook = sheet.getWorkbook();
    System.out.println("�e��Workbook�ɂ�" + 
          parentBook.getNumberOfSheets() +
          "���̃V�[�g������A����" + 
          (parentBook.getSheetIndex(sheet) + 1) +
          "�Ԗڂł��B");
  }
  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // �V�[�g��5������
    for (int i=0; i<5; i++) {
      workBook.createSheet();
    }
    // 3�Ԗڂ̃V�[�g�������ɃV�[�g�ŗL�������Ăяo��
    sheetProc(workBook.getSheetAt(2));

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
    // �����̎��s
    new GetParentBookTest().Run(args[0]);
    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
