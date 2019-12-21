import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * �n�C�p�[�����N�ݒ�̃e�X�g
 */
public class SetHyperLinkTest {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {

    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
		// CreationHelper�̎擾
		CreationHelper cHelper = workBook.getCreationHelper();
    // �V�[�g�̐��� 
    Sheet sheet = workBook.createSheet();
    // Row��Cell�̐���
    Row row = sheet.createRow(0);
    Cell cell = row.createCell(0);
    // Cell�ɕ�����ݒ�
    cell.setCellValue("POI�z�[���y�[�W");
		// �n�C�p�[�����N�̐���
		Hyperlink link = cHelper.createHyperlink(Hyperlink.LINK_URL);
		link.setAddress("http://poi.apache.org/");
		// �n�C�p�[�����N��URL�ݒ�
		cell.setHyperlink(link);
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./Book1.xls" : 
                      "./Book1.xlsx");
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
      System.out.println("�G���[�F���[�h���w�肵�ĉ������B");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }
    new SetHyperLinkTest().Run(args[0]);
  }
}
