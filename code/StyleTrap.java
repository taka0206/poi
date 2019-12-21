import java.io.*;
import org.apache.poi.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * �Z���X�^�C���̐ݒ�@��肠��o�[�W����
 */ 
public class StyleTrap {

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
    // �Z���X�^�C������
    CellStyle style = workBook.createCellStyle();
    for (int i=1; i<=30; i++) {
      // �Z���𐶐�
      Cell cell = row.createCell(i);
      cell.setCellValue("��");
      int mod = i % 3;
      switch (mod) {
        case 1:
          style.setAlignment(CellStyle.ALIGN_LEFT);
          break;
        case 2:
          style.setAlignment(CellStyle.ALIGN_CENTER);
          break;
        default:
          style.setAlignment(CellStyle.ALIGN_RIGHT);
          break;
      }
      cell.setCellStyle(style);
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
    new StyleTrap().Run(args[0]);
  }
}
