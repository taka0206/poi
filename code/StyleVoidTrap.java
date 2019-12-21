import java.io.*;
import org.apache.poi.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;
/**
 * �Z���X�^�C���̐ݒ� �C���o�[�W����
 */ 
public class StyleVoidTrap {
  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // ���[�N�V�[�g����
    Sheet sheet = workBook.createSheet("Sheet1");
    // Row�𐶐�
    Row row = sheet.createRow(5);
    // �Z���X�^�C����3����
    for (int i=0; i<3; i++) {
      workBook.createCellStyle();
    }
    // �쐬�����Z���ɏ�����ݒ肷��B
    CellStyle style1 = workBook.getCellStyleAt((short)21);
    style1.setAlignment(CellStyle.ALIGN_LEFT);
    CellStyle style2 = workBook.getCellStyleAt((short)22);
    style2.setAlignment(CellStyle.ALIGN_CENTER);
    CellStyle style3 = workBook.getCellStyleAt((short)23);
    style3.setAlignment(CellStyle.ALIGN_RIGHT);

    for (int i=1; i<=30; i++) {
      // �Z���𐶐�
      Cell cell = row.createCell(i);
      cell.setCellValue("��");
      int mod = i % 3;
      switch (mod) {
        case 1:
          cell.setCellStyle(style1);
          break;
        case 2:
          cell.setCellStyle(style2);
          break;
        default:
          cell.setCellStyle(style3);
          break;
      }
    }
   // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? this.getClass().getName() + "_Book1.xls" : 
                      this.getClass().getName() + "_Book1.xlsx");
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
    new StyleVoidTrap().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
