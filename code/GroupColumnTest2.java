import java.io.*;
import org.apache.poi.ss.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
/**
 * ��̃O���[�v���e�X�g
 */
class GroupColumnTest2 {
  /** �����̎��s
   * @param ���[�h
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    // ���[�N�u�b�N��ǂݍ���
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./group.xls" : "./group.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + e.toString());
      return;
    }
    // ����\�V�[�g�̎擾
    Sheet sheet = workBook.getSheetAt(0);
/*
    // �O���[�v������
    // 2�񂩂�4����O���[�v��
    sheet.groupColumn(1,3);
    // 7�񂩂�9����O���[�v��
    sheet.groupColumn(6,8);
*/
/*
    // �O���[�v��������
    // 2�񂩂�4����O���[�v����
    sheet.ungroupColumn(1,3);
    // 7�񂩂�9����O���[�v����
    sheet.ungroupColumn(6,8);
*/
    // �O���[�v���ꊇ����
		int lastCol = mode.equals("2003") ? SpreadsheetVersion.EXCEL97.getLastColumnIndex()
											: SpreadsheetVersion.EXCEL2007.getLastColumnIndex();
		// �O���[�v��W�J����
		//sheet.setColumnGroupCollapsed(4, false);
		//sheet.setColumnGroupCollapsed(9, false);
		// �O���[�v������������
		sheet.ungroupColumn(0, lastCol);
		for (int i=0;i<lastCol; i++) {
			sheet.setColumnWidth(i, sheet.getColumnWidth(i));
		}
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./Book1.xls" : 
                      "./Book1.xlsx");
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
    new GroupColumnTest2().Run(args[0]);
  }
}
