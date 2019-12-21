import java.io.*;
import org.apache.poi.ss.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.*;
/**
 * �o�b�`�t�@�C���쐬2
 */
class LancheBat2 {

  protected String _mode;   // ���샂�[�h
  protected Workbook _workBook; // �����`���[���[�N�u�b�N�̃C���X�^���X
  protected int _classPos;    // �N���X���̌��ʒu
  protected String[] _breakKeys; // �L�[�u���[�N���o�ޔ�̈�
  protected int[] _breakPos; // �L�[�u���[�N�s�ԍ��ޔ�̈�
  /** 
   * ��������
   */
  protected boolean init() {
    FileInputStream fis = null;
    // ���[�N�u�b�N��ǂݍ���
    _workBook = null;
    try {
      fis = new FileInputStream("./input/SampleLauncherORG.xls");
      _workBook = new HSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + e.toString());
      return false;
    }
    return true;
  }
  /** 
   * �o�b�`�t�@�C���쐬����
   */
  protected void buildBat() {
    // �f�[�^�V�[�g�̎擾
    Sheet dSheet = _workBook.getSheet("�f�[�^�V�[�g");
    // 2�s�ڂ���ŏI�s�܂ŏ���
    for (int i=2; i<dSheet.getLastRowNum(); i++) {
      Row row = dSheet.getRow(i);
      String className = row.getCell(4).getStringCellValue(); // �N���X��
      if (className.equals("�Ȃ�") == false) {
        if (row.getCell(8).getBooleanCellValue() == false) {
          // ������͂̂�͏������Ȃ��B
          if (row.getCell(7).getBooleanCellValue() == false) {
            // Book�𐶐������͏������Ȃ��B
            if (row.getCell(5).getBooleanCellValue() == false) {
              // �r���h�R�}���h�o��
              System.out.println("javac " + className + ".java");
              // ���s�R�}���h�o��(2003)
              System.out.println("java " + className + " 2003");
              if (row.getCell(6).getBooleanCellValue() == true) {
                // ���s�R�}���h�o��(2007)
                System.out.println("java " + className + " 2007");
              }
            }
          }
        }
      }
    }
  }
  /** �����̎��s
   * @param ���[�h
   */
  public void Run() {

    // ��������
    if (init() == false) {
      return;
    }
    // �o�b�`�t�@�C���쐬
    buildBat();
  }
  /** �G���g���[�|�C���g */
  public static void main(String[] args) {

    new LancheBat2().Run();
  }
}
