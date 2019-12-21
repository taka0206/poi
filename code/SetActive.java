import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

/**
 * �A�N�e�B�u�Z���A�A�N�e�B�u�V�[�g�ݒ�N���X
 */
class SetActive {
  /**
   * �A�N�e�B�u�Z���A�A�N�e�B�u�V�[�g�̐ݒ菈�����C��
   * @param fName ���[�N�u�b�N�t�@�C����
   */
  public void setActiveMain(String fName) {
    // ���[�N�u�b�N��ǂݍ���
    Workbook workBook = null;
    try {
      workBook = WorkbookFactory.create(
                new FileInputStream(fName));
    }
    catch(Exception e) {
      System.out.println(e.toString());
      return;
    }
    // �܂��S�V�[�g�̑I����Ԃ�����
    for (int i = 0; 
        i < workBook.getNumberOfSheets(); i++) {
      workBook.getSheetAt(i).setSelected(false);
    }
    // ���[�N�u�b�N�ɑ��݂���V�[�g�����ɏ������A
    // A1�Z�����A�N�e�B�u�ɂ���B
    for (int i=0;
          i<workBook.getNumberOfSheets(); i++) {
      Sheet sheet = workBook.getSheetAt(i);
      // 1�s�ڂ�Row���擾
      Row row = sheet.getRow(0);
      if (row == null) {
        // 1�s�ڂ�Row�����݂��Ȃ��ꍇ�B
        Row nrow = sheet.createRow(0);
        Cell cell = nrow.createCell(0);
        // A1�Z�����A�N�e�B�u�ɁB
        cell.setAsActiveCell();
      }
      else {
        // 1�s�ڂ�Row�����݂���ꍇ�B
        Cell cell = row.getCell(
              0, Row.CREATE_NULL_AS_BLANK);
        // A1�Z�����A�N�e�B�u�ɁB
        cell.setAsActiveCell();
      }
    }
    // �Ō�ɑ�1�V�[�g���A�N�e�B�u�ɁB
    workBook.setActiveSheet(0);
    workBook.getSheetAt(0).setSelected(true);
    // ���[�N�u�b�N�������o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream(fName);
      workBook.write(out);
    }catch(IOException e){
      System.out.println(e.toString());
      return;
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println(e.toString());
        return;
      }
    }
    System.out.println(fName + "���������܂����B");
  }
  /** �����̎��s
   * @param path ��������t�H���_(�f�B���N�g��)��
   */
  public void Run(String path) {
    File dir = new File(path);
    if (!dir.exists()) {
      System.out.println("�w�肳�ꂽ�p�X�͑��݂��܂���B");
      return;
    }
    File[] files = dir.listFiles();
    for (int i = 0; i < files.length; i++) {
      File file = files[i];
      if (file.isFile() && file.canRead()) {
        if (file.getPath().toLowerCase().endsWith(".xls") ||
            file.getPath().toLowerCase().endsWith(".xlsx")) {
          setActiveMain(file.getPath());
        }
      }
    }
  }
  /** �G���g���[�|�C���g */
  public static void main(String[] args) {
    if (args.length != 1) {
      System.out.println(
        "�G���[�F�g����-> SetActive �t�H���_��");
      return;
    }
    new SetActive().Run(args[0]);
  }
}
