import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * ����V�[�g�쐬�c�[��
 */
public class MakeUriageSheet {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N��ǂݍ���
		FileInputStream fis = null;
		Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/uriage.xls" : "./input/uriage.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + e.toString());
      return;
    }
    // �V�[�g�̎擾
    Sheet sheet = workBook.getSheetAt(0);
		Random rand = new Random();
		// 4�s�ڂ��珇�Ԃɏ���
		for (int i=3; i<=sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			// 2�J�����ڂ��珈��
			for (int j=1; j<13; j++) {
				Cell cell = row.getCell(j);
				cell.setCellValue(rand.nextInt(3000));
			}
		}
		if (mode.equals("2003")) {
			sheet.setForceFormulaRecalculation(true);
		}
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./uriage2.xls" : 
                      "./uriage2.xlsx");
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
    new MakeUriageSheet().Run(args[0]);
  }
}
