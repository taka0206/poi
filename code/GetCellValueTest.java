import java.io.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
/**
 * Cell�̒l�擾�e�X�g
 */
class GetCellValueTest {
  /** �����̎��s
   * @param ���[�h
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    // ���[�N�u�b�N��ǂݍ���
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/GetCellValue.xls" : "./input/GetCellValue.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + e.toString());
      return;
    }
    // 1�Ԗڂ̃V�[�g�̎擾
    Sheet sheet = workBook.getSheetAt(0);
    int j;
    for(int i=1; i<13; i++) {
      j = 3;
      Row row = sheet.getRow(i);
      Cell cell = row.getCell(1, row.CREATE_NULL_AS_BLANK);
      Cell cellDst = null;

      cellDst = row.getCell(j++, row.CREATE_NULL_AS_BLANK);
      // getStringCellvalue
      try {
        String s = cell.getStringCellValue();
        if (s==null) {
          cellDst.setCellValue("null");
        }
        else if (s.equals("")) {
          cellDst.setCellValue("�󕶎�");
        }
        else {
          cellDst.setCellValue("��");
        }
      }
      catch(Exception e) {
        System.out.println(e.toString());
        cellDst.setCellValue("�~");
      }
      // getRitchStringCellValue
      cellDst = row.getCell(j++, row.CREATE_NULL_AS_BLANK);
      try {
        RichTextString rs = cell.getRichStringCellValue();
        cellDst.setCellValue("��");
      }
      catch(Exception e) {
        System.out.println(e.toString());
        cellDst.setCellValue("�~");
      }
      // getDateCellValue
      cellDst = row.getCell(j++, row.CREATE_NULL_AS_BLANK);
      try {
        if(DateUtil.isCellDateFormatted(cell)) {
          Date dt = cell.getDateCellValue();
          cellDst.setCellValue("���t" + dt.toString());
        }
        else {
          cellDst.setCellValue("���t�łȂ�");
        }
      }
      catch(Exception e) {
        System.out.println(e.toString());
        cellDst.setCellValue("�~");
      }
      // getNumericCellValue
      cellDst = row.getCell(j++, row.CREATE_NULL_AS_BLANK);
      try {
        double db = cell.getNumericCellValue();
        cellDst.setCellValue("��(" + db + ")");
      }
      catch(Exception e) {
        System.out.println(e.toString());
        cellDst.setCellValue("�~");
      }
      // getBooleanCellValue
      cellDst = row.getCell(j++, row.CREATE_NULL_AS_BLANK);
      try {
        boolean b = cell.getBooleanCellValue();
        cellDst.setCellValue("��(" + b + ")" );
      }
      catch(Exception e) {
        System.out.println(e.toString());
        cellDst.setCellValue("�~");
      }
      // getCellFormula
      cellDst = row.getCell(j++, row.CREATE_NULL_AS_BLANK);
      try {
        String cf = cell.getCellFormula();
        cellDst.setCellValue("��");
      }
      catch(Exception e) {
        System.out.println(e.toString());
        cellDst.setCellValue("�~");
      }
      // getHyperlink() 
      cellDst = row.getCell(j++, row.CREATE_NULL_AS_BLANK);
      try {
        Hyperlink hl = cell.getHyperlink();
        if (hl != null) {
          cellDst.setCellValue("��");
        }
        else {
          cellDst.setCellValue("null");
        }
      }
      catch(Exception e) {
        System.out.println(e.toString());
        cellDst.setCellValue("�~");
      }

      // getErrorCellValue
      cellDst = row.getCell(j++, row.CREATE_NULL_AS_BLANK);
      try {
        byte bt = cell.getErrorCellValue();
        cellDst.setCellValue("��");
      }
      catch(Exception e) {
        System.out.println(e.toString());
        cellDst.setCellValue("�~");
      }
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
    new GetCellValueTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
