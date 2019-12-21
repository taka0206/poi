import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * �n�C�p�[�����N�����̃e�X�g
 */
public class RemoveLinkTest {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    FileInputStream fis = null;

    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // �V�[�g�̐��� Sheet0
    Sheet sheet = workBook.createSheet();
    // Sheet1
    Sheet sheet2 = workBook.createSheet();
    // Row��Cell�̐���
    for (int i=0; i<5; i++){
      sheet.createRow(i);
    }
    // Sheet1��Row��Cell�𐶐����l�ݒ�
    sheet2.createRow(0).createCell(0).setCellValue("�����N��");
    
    CreationHelper cHelper = workBook.getCreationHelper();
    // �n�C�p�[�����N�ݒ�
    // URL
    Cell cellUrl = sheet.getRow(0).createCell(0);
    Hyperlink linkURL = cHelper.createHyperlink(Hyperlink.LINK_URL);
    linkURL.setAddress("http://poi.apache.org/");
    //linkURL.setAddress("");
    cellUrl.setCellValue("�|�|�C�b��POI");
    cellUrl.setHyperlink(linkURL);
    // Document
    Cell cellDoc = sheet.getRow(1).createCell(0);
    Hyperlink linkDoc = cHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
    linkDoc.setAddress("SHeet1!A1");
    //linkDoc.setAddress("");
    cellDoc.setCellValue("Sheet1��");
    cellDoc.setHyperlink(linkDoc);
    // Mail
    Cell cellEMail = sheet.getRow(2).createCell(0);
    Hyperlink linkMail = cHelper.createHyperlink(Hyperlink.LINK_EMAIL);
    linkMail.setAddress("mailto:impl_person@yahoo.co.jp");
    //linkMail.setAddress("");
    cellEMail.setCellValue("�������͂�����");
    cellEMail.setHyperlink(linkMail);
    // File
    Cell cellFile = sheet.getRow(3).createCell(0);
    Hyperlink linkFile = cHelper.createHyperlink(Hyperlink.LINK_FILE);
    linkFile.setAddress("Book1.xlsx");
    //linkFile.setAddress("");
    cellFile.setCellValue("�ʂ̃u�b�N");
    cellFile.setHyperlink(linkFile);
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
    new RemoveLinkTest().Run(args[0]);
  }
}
