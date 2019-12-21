import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * ���t�ݒ�̃e�X�g
 */
public class SetDateTest {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {

    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // �V�[�g�̐��� 
    Sheet sheet = workBook.createSheet();
    // CellStyle����
    CellStyle style = workBook.createCellStyle();
    // ���t�����̐ݒ�
    DataFormat df = workBook.createDataFormat();
    style.setDataFormat(df.getFormat("yyyy�Nmm��dd��"));
    // 1��ڂ͓��t�Ȃ̂ŁA�����ňꊇ�����ݒ���\
    //sheet.setDefaultColumnStyle(0, style);
    // Row��5�s�ACell��1�񂸂���
    for (int i=0; i<5; i++) {
      sheet.createRow(i).createCell(0);
    }
    // ���t�ݒ� ����(Date)
    sheet.getRow(0).getCell(0).setCellValue(new Date());
    sheet.getRow(0).getCell(0).setCellStyle(style);
    // ���t�ݒ� ����(Calender)
    sheet.getRow(1).getCell(0).setCellValue(
        Calendar.getInstance());
    sheet.getRow(1).getCell(0).setCellStyle(style);
    sheet.autoSizeColumn(0); // �񕝎����ݒ�ɂ��Ă���
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? this.getClass().getName() + "_Book1.xls" : 
                      this.getClass().getName() + "_Book1.xlsx");
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
    // �����̎��s
    new SetDateTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
