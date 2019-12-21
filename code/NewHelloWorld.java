import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * Java+POI�Ő��E�Ɉ��A����v���O����
 * ���x�[�X�ɁARow�����ݒ胁�\�b�h�Ƃ̊֘A�𒲍��B
 *
 */
public class NewHelloWorld{
  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? 
            new HSSFWorkbook() : 
            new XSSFWorkbook();
    // ���[�N�V�[�g�̐���
    Sheet sheet = workBook.createSheet("HelloWorld");
    // Row�̐���
    Row row = sheet.createRow(1);
    // cell�̐���
    Cell cell = row.createCell(0);
    // cell�X�^�C���̐���
    CellStyle st = workBook.createCellStyle();
    // �t�H���g�̐���
    Font fnt = workBook.createFont();
    fnt.setFontName("�l�r ����");
    fnt.setFontHeightInPoints((short)48);
    fnt.setColor((short)HSSFColor.AQUA.index);
    // cell�X�^�C���Ƀt�H���g�ݒ�
    st.setFont(fnt);
    // cell�ɃX�^�C���ݒ�
    cell.setCellStyle(st);
    // cell�ɒl�ݒ�
    cell.setCellValue("Hello World��");

    // �l�r ����48�|�C���g��55.5�s�N�Z���ɂȂ�̂ŁA
    // �����菬�����l(����)��Row�̍�����ݒ肵�Ă݂�B
    row.setHeightInPoints((float)25.25);

    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( 
        mode.equals("2003") ? 
        this.getClass().getName() + "_Book1.xls" : 
        this.getClass().getName() + "_Book1.xlsx");
      workBook.write(out);
    }catch(IOException e){
      System.out.println(
        "�u�b�N�̏������݂Ɏ��s���܂����B\n" + 
        e.toString());
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println(
          "�u�b�N�̏������݂Ɏ��s���܂����B\n" + 
          e.toString());
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
    else if ( !args[0].equals("2003") && 
              !args[0].equals("2007") ) {
      System.out.println(
        "�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }
    // �����̎��s
    new NewHelloWorld().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
