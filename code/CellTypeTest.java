import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
/**
 * CellType�̎擾�e�X�g
 */
class CellTypeTest {
  /** �����̎��s
   * @param ���[�h
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    // ���[�N�u�b�N��ǂݍ���
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/CellType.xls" : "./input/CellType.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + e.toString());
      return;
    }
    String typeString[] = new String[] {
      "CELL_TYPE_NUMERIC", 
      "CELL_TYPE_STRING",
      "CELL_TYPE_FOMULA",
      "CELL_TYPE_BLANK",
      "CELL_TYPE_BOOLEAN",
      "CELL_TYPE_ERR"};
    // 1�Ԗڂ̃V�[�g�̎擾
    Sheet sheet = workBook.getSheetAt(0);
    // B2Cell����B14Cell�܂ŏ��Ԃ�CellType�𔻒�
    for (int i=1; i<14; i++) {
      Row row = sheet.getRow(i);
      Cell cellDst = row.getCell(2,
        row.CREATE_NULL_AS_BLANK);
      System.out.println(i + ":" + 
        typeString[row.getCell(
          1, row.CREATE_NULL_AS_BLANK).getCellType()]);
      cellDst.setCellValue(typeString[row.getCell(
          1, row.CREATE_NULL_AS_BLANK).getCellType()]);
    }
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
      System.out.println("�G���[�F���[�h���w�肵�Ă��������B");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }
    // �����̎��s
    new CellTypeTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
