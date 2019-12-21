import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * �O�i�F�E�w�i�F�E�p�^�[���̃e�X�g
 */
public class SetCellColorTest {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {

    // �h��ׂ��p�^�[����
    String[] fillPatNames = {
                             "NO_FILL"
                            ,"SOLID_FOREGROUND"
                            ,"FINE_DOTS"
                            ,"ALT_BARS"
                            ,"SPARSE_DOTS"
                            ,"THICK_HORZ_BANDS"
                            ,"THICK_VERT_BANDS"
                            ,"THICK_BACKWARD_DIAG"
                            ,"THICK_FORWARD_DIAG"
                            ,"BIG_SPOTS"
                            ,"BRICKS"
                            ,"THIN_HORZ_BANDS"
                            ,"THIN_VERT_BANDS"
                            ,"THIN_BACKWARD_DIAG"
                            ,"THIN_FORWARD_DIAG"
                            ,"SQUARES"
                            ,"DIAMONDS"
                            ,"LESS_DOTS"
                            ,"LEAST_DOTS"
    };
    // �h��ׂ��p�^�[���l
    short[] fillPatValues = {
                             CellStyle.NO_FILL
                            ,CellStyle.SOLID_FOREGROUND
                            ,CellStyle.FINE_DOTS
                            ,CellStyle.ALT_BARS
                            ,CellStyle.SPARSE_DOTS
                            ,CellStyle.THICK_HORZ_BANDS
                            ,CellStyle.THICK_VERT_BANDS
                            ,CellStyle.THICK_BACKWARD_DIAG
                            ,CellStyle.THICK_FORWARD_DIAG
                            ,CellStyle.BIG_SPOTS
                            ,CellStyle.BRICKS
                            ,CellStyle.THIN_HORZ_BANDS
                            ,CellStyle.THIN_VERT_BANDS
                            ,CellStyle.THIN_BACKWARD_DIAG
                            ,CellStyle.THIN_FORWARD_DIAG
                            ,CellStyle.SQUARES
                            ,CellStyle.DIAMONDS
                            ,CellStyle.LESS_DOTS
                            ,CellStyle.LEAST_DOTS
    };

    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // �V�[�g�̐��� 
    Sheet sheet = workBook.createSheet("Cell�h��ׂ��p�^�[��");
    
    for (int i=0; i<19;i++) {
      Row row = sheet.createRow(i);
      row.createCell(0).setCellValue(fillPatNames[i]);
      row.createCell(1).setCellValue("(" + fillPatValues[i] +")");
      Cell cell = row.createCell(2);
      CellStyle style = workBook.createCellStyle();
      style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
      style.setFillBackgroundColor(IndexedColors.WHITE.getIndex());
      style.setFillPattern(fillPatValues[i]);
      cell.setCellStyle(style);
    }
    sheet.autoSizeColumn(0);  // 1��ڂ��������ݒ��
    sheet.setZoom(2,1);       // 200%�Ɋg��

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
    new SetCellColorTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
