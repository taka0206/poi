import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
/**
 * �V�[�g�ی샏�[�N�u�b�N�̓ǂݍ��݂���я�������
 */
class ReadProtectSheet {
  /** �����̎��s
   * @param ���[�h
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    // ���[�N�u�b�N��ǂݍ���
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/protect.xls" : "./input/protect.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + e.toString());
      //return;
    }
    // ��Ԗڂ̃V�[�g�̎擾
    Sheet sheet = workBook.getSheetAt(0);
    Row row = sheet.getRow(0);
    if (row == null) {
      Row rown = sheet.createRow(0);
      rown.createCell(0).setCellValue("UPDATE");
    }
    else {
      Cell cell = row.getCell(
                    0, Row.CREATE_NULL_AS_BLANK);
      cell.setCellValue("UPDATE");
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
      System.out.println("�G���[�F���[�h���w�肵�Ă��������B");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }
    // �����̎��s
    new ReadProtectSheet().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
