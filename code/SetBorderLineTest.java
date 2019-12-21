import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

/**
 * �r���̃e�X�g
 */
public class SetBorderLineTest {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {

    // ���햼
    String[] linePatNames = {
                             "BORDER_NONE"
                            ,"BORDER_THIN"
                            ,"BORDER_MEDIUM"
                            ,"BORDER_DASHED"
                            ,"BORDER_HAIR"
                            ,"BORDER_THICK"
                            ,"BORDER_DOUBLE"
                            ,"BORDER_DOTTED"
                            ,"BORDER_MEDIUM_DASHED"
                            ,"BORDER_DASH_DOT"
                            ,"BORDER_MEDIUM_DASH_DOT"
                            ,"BORDER_DASH_DOT_DOT"
                            ,"BORDER_MEDIUM_DASH_DOT_DOT"
                            ,"BORDER_SLANTED_DASH_DOT"
    };
    // ����l
    short[] linePatValues = {
                             CellStyle.BORDER_NONE
                            ,CellStyle.BORDER_THIN
                            ,CellStyle.BORDER_MEDIUM
                            ,CellStyle.BORDER_DASHED
                            ,CellStyle.BORDER_HAIR
                            ,CellStyle.BORDER_THICK
                            ,CellStyle.BORDER_DOUBLE
                            ,CellStyle.BORDER_DOTTED
                            ,CellStyle.BORDER_MEDIUM_DASHED
                            ,CellStyle.BORDER_DASH_DOT
                            ,CellStyle.BORDER_MEDIUM_DASH_DOT
                            ,CellStyle.BORDER_DASH_DOT_DOT
                            ,CellStyle.BORDER_MEDIUM_DASH_DOT_DOT
                            ,CellStyle.BORDER_SLANTED_DASH_DOT
    };

    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // �V�[�g�̐��� 
    Sheet sheet = workBook.createSheet("�r��");
    int rowNo = 0;
    for (int i=0; i<14; i++) {
      Row row = sheet.createRow(rowNo);
      Cell cell0 = row.createCell(0);
      cell0.setCellValue(linePatNames[i]);
      Cell cell1 = row.createCell(1);
      cell1.setCellValue("(" + linePatValues[i] + ")");
      // �Z������������B
      sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo + 1, 0,0));
      sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo + 1, 1,1));
      // �����Z���̃X�^�C��
      CellStyle styleM = workBook.createCellStyle();
      styleM.setAlignment(CellStyle.ALIGN_RIGHT);
      styleM.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
      cell0.setCellStyle(styleM);
      cell1.setCellStyle(styleM);
      Cell cell = sheet.createRow(rowNo+1).createCell(2);
      // CellStyle����
      CellStyle styleLine = workBook.createCellStyle();
      // �����ݒ�
      styleLine.setBorderTop(linePatValues[i]);
      cell.setCellStyle(styleLine);
      rowNo += 2;
    }
    sheet.autoSizeColumn(0,true);  // 1��ڂ��������ݒ��(�}�[�W�Ώ�)
    sheet.autoSizeColumn(1,true);  // 2��ڂ��������ݒ��(�}�[�W�Ώ�)
    sheet.setColumnWidth(2, 12800); // 3��ڂ��L��
    sheet.setDisplayGridlines(false); // �r�������Ղ��悤�V�[�g�g���������B

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
    new SetBorderLineTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
