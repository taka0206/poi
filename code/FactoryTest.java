import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

/**
 * WorkbookFactory�̃e�X�g
 */
class FactoryTest {
  /** �����̎��s
   * @param fName �ǂݍ��ރt�@�C����
   */
  public void Run(String fName) {
    // ���[�N�u�b�N��ǂݍ���
    Workbook workBook = null;
    try {
      workBook = WorkbookFactory.create(
          new FileInputStream(fName));
    }
    catch(Exception e) {
      System.out.println(e.toString());
    }
    // Excel�h�L�������g�`���𔻒肷��B
    if (workBook instanceof HSSFWorkbook) {
      System.out.println("Excel2003�ȑO�̌`���ł��B");
    }
    else if(workBook instanceof XSSFWorkbook) {
      System.out.println("Excel2007�ȍ~�̌`���ł��B");
    }
    else {
      System.out.println("�s���Ȍ`���ł��B");
    }
  }
  /** �G���g���[�|�C���g */
  public static void main(String[] args) {
    String inputValue;
    // �ǂݍ��ݑΏۃt�@�C������
    while (true) {
      System.out.print("�ǂݍ���Excel�t�@�C��������͂��Ă�������(�t���p�X)�B���~(x) ->");
      BufferedReader buf =
              new BufferedReader(
                     new InputStreamReader(System.in),1);
      try {
        inputValue = buf.readLine().toLowerCase();
      }
      catch (Exception e)
      {
        System.out.println("�t�@�C�������͂ŃG���[���������܂����B" + e.toString());
        return;
      }
      if (inputValue.equals("x")) {
        return;
      }
      if (inputValue.length() != 0) {
        break;
      }
    }
    // �����̎��s
    new FactoryTest().Run(inputValue);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
