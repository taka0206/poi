import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

/**
 * �v�����^�[�ݒ�擾�e�X�g
 */
public class GetPrintSetupTest {
  /**
   * �p���T�C�Y������擾
   *@pSize �p���T�C�Y�ԍ�
   */
  protected String deCodePaparSize(short pSize) {
    switch (pSize) {
      case PrintSetup.A3_PAPERSIZE :
        return "A3 - 297x420 mm";
      case PrintSetup.A4_EXTRA_PAPERSIZE :
        return "A4 Extra - 9.27 x 12.69 in";
      case PrintSetup.A4_PAPERSIZE :
        return "A4 - 210x297 mm";
      case PrintSetup.A4_PLUS_PAPERSIZE :
        return "A4 Plus - 210x330 mm";
      case PrintSetup.A4_ROTATED_PAPERSIZE :
        return "A4 Rotated - 297x210 mm";
      case PrintSetup.A4_SMALL_PAPERSIZE :
        return "A4 Small - 210x297 mm";
      case PrintSetup.A4_TRANSVERSE_PAPERSIZE :
        return "A4 Transverse - 210x297 mm";
      case PrintSetup.A5_PAPERSIZE :
        return "A5 - 148x210 mm";
      case PrintSetup.B4_PAPERSIZE :
        return "B4 (JIS) 250x354 mm";
      case PrintSetup.B5_PAPERSIZE :
        return "B5 (JIS) 182x257 mm";
      case PrintSetup.ELEVEN_BY_SEVENTEEN_PAPERSIZE :
        return "11 x 17 in";
      case PrintSetup.ENVELOPE_10_PAPERSIZE :
        return "US Envelope #10 4 1/8 x 9 1/2";
      case PrintSetup.ENVELOPE_9_PAPERSIZE :
        return "US Envelope #9 3 7/8 x 8 7/8";
      case PrintSetup.ENVELOPE_C3_PAPERSIZE :
        return "Envelope C3 324x458 mm";
      case PrintSetup.ENVELOPE_C4_PAPERSIZE :
        return "Envelope C4 229x324 mm";
      case PrintSetup.ENVELOPE_C5_PAPERSIZE :
        return "Envelope C5";
      case PrintSetup.ENVELOPE_C6_PAPERSIZE :
        return "Envelope C6 114x162 mm";
      case PrintSetup.ENVELOPE_DL_PAPERSIZE :
        return "Envelope DL 110x220 mm";
      case PrintSetup.ENVELOPE_MONARCH_PAPERSIZE :
        return "Envelope Nonarch";
      case PrintSetup.EXECUTIVE_PAPERSIZE :
        return "US Executive 7 1/4 x 10 1/2 in";
      case PrintSetup.FOLIO8_PAPERSIZE :
        return "Folio 8 1/2 x 13 in";
      case PrintSetup.LEDGER_PAPERSIZE :
        return "US Ledger 17 x 11 in";
      case PrintSetup.LEGAL_PAPERSIZE :
        return "US Legal 8 1/2 x 14 in";
      case PrintSetup.LETTER_PAPERSIZE :
        return "US Letter 8 1/2 x 11 in";
      case PrintSetup.LETTER_ROTATED_PAPERSIZE :
        return "US Letter Rotated 11 x 8 1/2 in";
      case PrintSetup.LETTER_SMALL_PAGESIZE :
        return "US Letter Small 8 1/2 x 11 in";
      case PrintSetup.NOTE8_PAPERSIZE :
        return "US Note 8 1/2 x 11 in";
      case PrintSetup.QUARTO_PAPERSIZE :
        return "Quarto 215x275 mm";
      case PrintSetup.STATEMENT_PAPERSIZE :
        return "US Statement 5 1/2 x 8 1/2 in";
      case PrintSetup.TABLOID_PAPERSIZE :
        return "US Tabloid 11 x 17 in";
      case PrintSetup.TEN_BY_FOURTEEN_PAPERSIZE :
        return "10 x 14 in";
    }
    return "unknown";
  }
  /**
   * �v�����^�[�ݒ�o��
   *@psetup PrintSetup�̎Q��
   */
  protected void PrintPrintSettings(PrintSetup psetup) {
    System.out.println("�������                 : " + 
      psetup.getCopies());
    System.out.println("FitHeight                : " + 
      psetup.getFitHeight());
    System.out.println("FitWidth                 : " + 
      psetup.getFitWidth());
    System.out.println("�t�b�^�[�]��             : " + 
      psetup.getFooterMargin());
    System.out.println("�w�b�_�[�]��             : " + 
      psetup.getHeaderMargin());
    System.out.println("�����𑜓x               : " + 
      psetup.getHResolution());
    System.out.println("LeftToRight              : " + 
      psetup.getLeftToRight());
    System.out.println("�����h�X�P�[�v���[�h     : " + 
      psetup.getLandscape());
    System.out.println("�������[�h               : " + 
      psetup.getNoColor());
    System.out.println("NoOrientation            : " + 
      psetup.getNoOrientation());
    System.out.println("�Z�����R�����g������[�h : " + 
      psetup.getNotes());
    System.out.println("PageStart                : " + 
      psetup.getPageStart());
    System.out.println("�p���T�C�Y               : " + 
      deCodePaparSize(psetup.getPaperSize()));
    System.out.println("����{��                 : " + 
      psetup.getScale());
    System.out.println("�y�[�W�ԍ����           : " + 
      psetup.getUsePage());
    System.out.println("ValidSettings            : " + 
      psetup.getValidSettings());
    System.out.println("�����𑜓x               : " + 
      psetup.getVResolution());
  }
  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // �V�[�g�̐���
    Sheet sheet = workBook.createSheet();
    // �v�����^�[�ݒ�̎擾
    PrintSetup psetup = sheet.getPrintSetup(); 
    // �e�ݒ���o�͂���B
    PrintPrintSettings(psetup);
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
    new GetPrintSetupTest().Run(args[0]);
    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
