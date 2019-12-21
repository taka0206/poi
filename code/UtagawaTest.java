import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 

/**
 * �^�킵�������؂���N���X
 * �������Ȃ������s��Z���͂ǂ��Ȃ��Ă���̂��B
 */
class UtagawaTest {
  protected String _mode;

  /**
   * �u�b�N�̏����o������
   */
  protected void writeBook() {
    // ���[�N�u�b�N�̐���
    Workbook workBook = _mode.equals("2003") ? new HSSFWorkbook() : new XSSFWorkbook();
    // ���[�N�V�[�g�𐶐�
    Sheet sheet = workBook.createSheet("Sheet1");
    // Row��10�ACell��20����
    for (int i=0; i<10; i++) {
      Row row = sheet.createRow(i);
      for (int j=0; j<20; j++) {
        row.createCell(j);
      }
    }
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( _mode.equals("2003") ? this.getClass().getName() + "_Book1.xls" : 
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
    System.out.println("WriteBook done!");
  }
  /**
   * �u�b�N�̓ǂݍ���
   */
  protected void readBook() {
    FileInputStream fis = null;
    Workbook workBook = null;
    // ���[�N�u�b�N�̓ǂݍ���
    try {
      fis = new FileInputStream( _mode.equals("2003") ? 
              this.getClass().getName() + "_Book1.xls" : this.getClass().getName() + "_Book1.xlsx");
      workBook = _mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println(e.toString());
    }
    Sheet sheet = workBook.getSheetAt(0);
    // �C�ӂ̃Z���ɒl�ݒ�
    // �����s�ƃZ�������݂��Ȃ���΂����ŗ�����͂��ł���B
    try {
      sheet.getRow(1).getCell(0).setCellValue("�Z���͂��邩");
    }
    catch (Exception e) {
      System.out.println("����ς�Z���͂Ȃ������I�I" + e.toString());
      return;
    }
    System.out.println("�Z���͑��݂��܂����B");
  }
  /** �����̎��s
   * @param mode ���[�h
   * @param rw   �����o���A�ǂݍ���
   */
  public void Run(String mode, String rw) {

    _mode = mode;

    if (rw.equals("w")) {
      writeBook();
    }
    else {
      readBook();
    }
  }
  /** �G���g���[�|�C���g */
  public static void main(String[] args) {
    if (args.length == 0) {
      System.out.println("�G���[�F�g����-> CalcTest ���[�h rw�t���O");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }
    String inputValue;
    if (args.length == 1) {
      while (true) {
        System.out.print("Read(r)/write(w)�̂����ꂩ���w�肵�Ă��������B���~(X) ->");
        BufferedReader buf =
                new BufferedReader(
                       new InputStreamReader(System.in),1);
        try {
          inputValue = buf.readLine().toLowerCase();
        }
        catch (Exception e)
        {
          System.out.println("rw�t���O���͂ŃG���[���������܂����B" + e.toString());
          return;
        }
        if (inputValue.equals("x")) {
          return;
        } 
        if ( !inputValue.equals("r") && !inputValue.equals("w") ) {
          System.out.println("�G���[�Frw�t���O��r�܂���w���w�肵�ĉ������B");
        }
        else {
          break;
        }
      }
    }
    else {
      if ( !args[1].equals("r") && !args[1].equals("w") ) {
        System.out.println("�G���[�Frw�t���O��r�܂���w���w�肵�ĉ������B");
        return;
      }
      inputValue = args[1];
    }
    new UtagawaTest().Run(args[0], inputValue);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
