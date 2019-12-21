import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 

/**
 * ���Ɗ֐��Čv�Z�̃e�X�g
 */
class CalcTest {
  /** �����̎��s
   * @param ���[�h
   * @value1 �����l1
   * @value2 �����l2
   */
  public void Run(String mode, int value1, int value2) {
    FileInputStream fis = null;
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/calctest.xls" : "./input/calctest.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println(e.toString());
    }
    // �V�[�g�̎擾
    Sheet sheet = workBook.getSheetAt(0);
    // Row�̎擾
    Row row = sheet.getRow(1);
    // A2�Z���ɒl�ݒ�(Null�Ȃ�Cell�𐶐�)
    row.getCell(0, 
      Row.CREATE_NULL_AS_BLANK).setCellValue(value1);
    // C2�Z���ɒl�ݒ�(Null�Ȃ�Cell�𐶐�)
    row.getCell(2, 
      Row.CREATE_NULL_AS_BLANK).setCellValue(value2);
    // E2�Z���̌v�Z����ǂݏo���čĐݒ肷��B
    String fum = row.getCell(4).getCellFormula();
    row.getCell(4).setCellFormula(fum);
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
    int value1,value2;
    if (args.length != 1) {
      System.out.println("�G���[�F���[�h���w�肵�ĉ������B");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }
    while(true) {
      String buf;
      // ���l1�̓���
      System.out.print("�����l1�̓���(X�Œ��~) -> ");
      InputStreamReader isr = new InputStreamReader(System.in);
      BufferedReader br = new BufferedReader(isr);
      try {
        buf = br.readLine();
        if (buf.toUpperCase().equals("X")) {
          return;
        }
        try {
          value1 = Integer.parseInt(buf);
          break;
        }
        catch (NumberFormatException e){
          System.out.println("�G���[�F�����l1�ɂ͐������w�肵�ĉ������B");
        }
      }
      catch(Exception e){
      }
    }
    while(true) {
      String buf;
      // ���l1�̓���
      System.out.print("�����l2�̓���(X�Œ��~) -> ");
      InputStreamReader isr = new InputStreamReader(System.in);
      BufferedReader br = new BufferedReader(isr);
      try {
        buf = br.readLine();
        if (buf.toUpperCase().equals("X")) {
          return;
        }
        try {
          value2 = Integer.parseInt(buf);
          break;
        }
        catch (NumberFormatException e){
          System.out.println("�G���[�F�����l2�ɂ͐������w�肵�ĉ������B");
        }
      }
      catch(Exception e) {
      }
    }

    // �����̎��s
    new CalcTest().Run(args[0],value1,value2);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
