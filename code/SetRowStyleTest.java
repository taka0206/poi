import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * Row�P�ʂ̕W���Z���X�^�C���ݒ�e�X�g
 */
public class SetRowStyleTest {

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
    // �Z���X�^�C������
    CellStyle style = workBook.createCellStyle();
    // �l�r���� 11�|�C���g�̃t�H���g�𐶐�
    Font fnt = workBook.createFont();
    fnt.setFontName("�l�r ����");
    fnt.setFontHeightInPoints((short)11);
    // �Z���X�^�C���Ƀt�H���g��ݒ�
    style.setFont(fnt);
    // Row��5�s����
		for (int i=0; i<5; i++) {
    	sheet.createRow(i);
    }
		// 3�s�ڂ�����CellStyle�ݒ�
		//if (mode.equals("2003")) ((HSSFRow)sheet.getRow(2)).setRowStyle((HSSFCellStyle)style);
    // �eRow��Cell��10������������ݒ�
		for (int i=0; i<5; i++) {
    	for (int j=0; j<10; j++) {
      	sheet.getRow(i).createCell(j).setCellValue(i + "-" + j);
			}
		}
		// Row������W���Z���X�^�C�����擾
		if (mode.equals("2003")) {
			HSSFCellStyle rstyle = ((HSSFRow)sheet.getRow(2)).getRowStyle();
			if (rstyle != null) {
				HSSFFont rfnt = rstyle.getFont((HSSFWorkbook)workBook);
				System.out.println(rfnt.getFontName() + rfnt.getFontHeightInPoints() + "�|�C���g");
			}
			else {
				System.out.println("�W��CellStyle��null�ł�");
			}
		}

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
    new SetRowStyleTest().Run(args[0]);
  }
}
