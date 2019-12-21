import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * �����ݕی�ݒ�̃e�X�g
 */
public class WriteProtectTestHSSF {

  /** 
   * �����̎��s
   */
  public void Run() {
    // ���[�N�u�b�N�̐���
    HSSFWorkbook workBook = new HSSFWorkbook();
    // �V�[�g���� 
    HSSFSheet sheet = workBook.createSheet();
    // �s��10�s�Z����10�쐬���Ēl�ݒ�
    for (int i=0; i<10; i++) {
      HSSFRow row = sheet.createRow(i);
      for (int j=0; j<10; j++) {
        row.createCell(j).setCellValue(i+"-"+j);
      }
    }
		// �����ݕی��ݒ�
		workBook.writeProtectWorkbook("POI", "POIAPI");
		// ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( "./Book1.xls");
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

    new WriteProtectTestHSSF().Run();
  }
}
