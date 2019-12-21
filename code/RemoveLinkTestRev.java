import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * �n�C�p�[�����N�����̃e�X�g(���x���W��)
 */
public class RemoveLinkTestRev {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N��ǂݍ���
		FileInputStream fis = null;
		Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./poilink.xls" : "./poilink.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + e.toString());
      return;
    }
    // �V�[�g�̎擾
    Sheet sheet = workBook.getSheetAt(0);
		// 2�s�ڂ��珇�Ԃɏ���
		for (int i=1; i<sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			Cell cellOrg = row.getCell(1);	// B�Z�����擾
			// ����ޔ�
			String sVal = cellOrg.getStringCellValue();	// �l
			Comment com = cellOrg.getCellComment();			// �Z���R�����g
			CellStyle style = cellOrg.getCellStyle();		// �Z���X�^�C��
			int type = cellOrg.getCellType();						// �Z���^�C�v
			// ����Cell���폜
			row.removeCell(cellOrg);
			// �����ꏊ��Cell�𐶐�
			/*
			Cell cellNew = row.createCell(1);
			// Cell�̏���ݒ肵�Ȃ����B
			cellNew.setCellValue(sVal);
			cellNew.setCellComment(com);
			cellNew.setCellStyle(style);
			cellNew.setCellType(type);
			*/
		}
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./hlink.xls" : 
                      "./hlink.xlsx");
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
    new RemoveLinkTestRev().Run(args[0]);
  }
}
