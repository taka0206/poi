import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;

/**
 * Java+POI�Ő��E�Ɉ��A����v���O����
 *
 */
public class HelloWorld{
  public static void main(String[] args){

    // ���[�N�u�b�N�̐���
    HSSFWorkbook workBook = new HSSFWorkbook();
    // ���[�N�V�[�g�̐���
    HSSFSheet sheet = 
        workBook.createSheet("HelloWorld");

    // Row�̐���
    HSSFRow row = sheet.createRow(0);

    // cell�̐���
    HSSFCell cell = row.createCell((short)0);

    // cell�X�^�C���̐���
    HSSFCellStyle st = workBook.createCellStyle();

    // �t�H���g�̐���
    HSSFFont fnt = workBook.createFont();
    fnt.setFontName("�l�r ����");
    fnt.setFontHeightInPoints((short)48);
    fnt.setColor((short)HSSFColor.AQUA.index);

    // cell�X�^�C���Ƀt�H���g�ݒ�
    st.setFont(fnt);

    // cell�ɃX�^�C���ݒ�
    cell.setCellStyle(st);

    // cell�ɒl�ݒ�
    cell.setCellValue("Hello World��");

    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream(
              "HelloWorld_Book1.xls");
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

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
