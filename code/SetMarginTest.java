import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * ������]���ݒ�e�X�g
 */ 
public class SetMarginTest {
  /** 
   * cm - �C���`�ϊ�
   * @param cm ����(�Z���`���[�g��)
   */
  protected double getInch(double cm) {
    return cm * 0.3937;
  }
  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
 
    // ���[�N�V�[�g����
    Sheet sheet = workBook.createSheet();
    // ������̗]����ݒ�
    // �㉺1.5cm
    sheet.setMargin(Sheet.TopMargin, getInch(1.5));
    sheet.setMargin(Sheet.BottomMargin, getInch(1.5));
    // ���E2cm
    sheet.setMargin(Sheet.LeftMargin, getInch(2.0));
    sheet.setMargin(Sheet.RightMargin, getInch(2.0));
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./" + this.getClass().getName() + "_Book1.xls" : 
                      "./" + this.getClass().getName() + "_Book1.xlsx");
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
    // �����̎��s
    new SetMarginTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}