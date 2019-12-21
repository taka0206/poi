import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * �e��Row�̂��̂܂��e�̃V�[�g�擾�e�X�g
 */
public class GetParentSheetByCellTest {

  /**
   * Cell�ŗL�̏���
   *@param cell �Z���̎Q��
   */
  public void cellProc(Cell cell) {
    Sheet pSheet = cell.getSheet();
    System.out.println("�e��Sheet��" + 
                pSheet.getSheetName() + "�ł��B");
    System.out.println(pSheet.getSheetName() + 
                "�ɂ�Row��" + 
                pSheet.getLastRowNum() + "�s����A");
    System.out.println("����" + 
                cell.getRowIndex() +
                "�s�ڂ�Row��" + 
                cell.getColumnIndex() + 
                "�߂̃Z���ł��B");
  }
  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    Random rand = new Random(); // ���������̏���
    int lim;
    // �V�[�g��5���A�e�V�[�g��Row���ő�10�sCell���ő�20����
    for (int i=0; i<5; i++) {
      Sheet sheet = workBook.createSheet();
      lim = rand.nextInt(10) + 1;
      for (int j=0; j<lim; j++) {
        Row row = sheet.createRow(j);
        lim = rand.nextInt(20) + 1;
        for (int k=0; k<lim; k++) {
          row.createCell(k).setCellValue(i + "-" + j + "-" + k);
        }
      }
    }
    // �m�F�p��Workbook���o��
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
        return;
      }
    }
    // ��������Cell������肷��B
    Sheet curSheet = workBook.getSheetAt(rand.nextInt(5));
    if (curSheet == null) {
      System.out.println("Sheet�擾���s");
      return;
    }
    Row curRow = curSheet.getRow(rand.nextInt(curSheet.getLastRowNum()));
    if (curRow == null) {
      System.out.println("Row�擾���s");
      return;
    }
    Cell curCell = curRow.getCell(rand.nextInt(curRow.getLastCellNum()));
    // Cell�ŗL�����Ăяo��
    cellProc(curCell);
  }
  /** �G���g���[�|�C���g */
  public static void main(String[] args) {
    if (args.length != 1) {
      System.out.println("�G���[�F���[�h���w�肵�ĉ������B");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }
    // �����̎��s
    new GetParentSheetByCellTest().Run(args[0]);
    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
