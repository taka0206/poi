import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 
import org.apache.poi.hssf.record.crypto.*;
/**
 * �v�Z���ݒ�̃e�X�g
 */
class setCellFormulaTest {
  /** �����̎��s
   * @param ���[�h
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    // ���[�N�u�b�N��ǂݍ���
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/setCellFormula_in.xls" : 
                  "./input/setCellFormula_in.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + e.toString());
      return;
    }
    // ���ϓ_�ҏW�p�������������Ă����B
    DataFormat df = workBook.createDataFormat();
    // 0�Ԗڂ�sheet���擾
    Sheet sheet = workBook.getSheetAt(0);
    // �l�ʍ��v���_�ƕ��όv�Z���̐ݒ�
    for (int i=3; i<23; i++) {
      sheet.getRow(i).getCell(6).setCellFormula(
        "SUM(C"+ (i+1) + ":F" + (i+1) + ")");
      sheet.getRow(i).getCell(7).setCellFormula(
        "AVERAGE(C"+ (i+1) + ":F" + (i+1) + ")");
      // CellStyle���擾���A������ݒ�
      CellStyle style = 
        sheet.getRow(i).getCell(7).getCellStyle();
      style.setDataFormat(df.getFormat("0.0"));
    }
    String colChr[] = {"A","B","C","D","E","F"};
    // �Ȗڕʕ��όv�Z����ݒ�
    Row row = sheet.getRow(23);
    for (int i=2; i<6; i++) {
      row.getCell(i).setCellFormula(
        "AVERAGE(" + colChr[i] + "4:" + 
        colChr[i] + "23)");
      CellStyle style = row.getCell(i).getCellStyle();
      style.setDataFormat(df.getFormat("0.0"));
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
    new setCellFormulaTest().Run(args[0]);
    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
