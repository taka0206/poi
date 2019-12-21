import javax.jws.*;
import java.io.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
/**
 * POI���b�p�[Web�T�[�r�X
 */
@WebService(targetNamespace="http://example.org")

public class CSWrapper {
  private Workbook _workBook = null;
  private Sheet _sheet = null;
  private Font _font = null;
  private CellStyle _style = null;
  private Row _row = null;
  private Cell _cell = null;
  /** �R���X�g���N�^ */
  public CSWrapper(){
  }
  /** 
   * ���[�N�V�[�g�A�s�A�Z������
   *@param mode ���샂�[�h 2003 or 2007
   *@param sName ���[�N�V�[�g�̖��O
   */ 
  public boolean createWorkSheetAndRowAndCell(
    @WebParam(name="mode") String mode,
    @WebParam(name="sName") String sName) {
    if ( !mode.equals("2003") && !mode.equals("2007")) {
      return false;
    }
    if (mode.equals("2003")) {
      _workBook = new HSSFWorkbook();
    }
    else {
      _workBook = new XSSFWorkbook();
    }
    _font = _workBook.createFont();
    _style = _workBook.createCellStyle();
    _sheet = _workBook.createSheet(sName);
    _row = _sheet.createRow(0);
    _cell = _row.createCell((short)0);
    return true;
  }
  /**
   * �t�H���g�̎w��ƃX�^�C���ݒ�
   *@param fontName �t�H���g��
   *@param po �����̃|�C���g
   *@param col �����F
   */
  public void setFontAndStyle(@WebParam(name="fontName") String fontName, 
                        @WebParam(name="po") int po, 
                        @WebParam(name="col") int col) {
    _font.setFontName(fontName);
    _font.setFontHeightInPoints((short)po);
    _font.setColor((short)col);
    _style.setFont(_font);
    _cell.setCellStyle(_style);
  }
  /**
   * �Z���ɕ����ݒ�
   *@param sVal �Z���ɐݒ肵��������
   */
  public void setCellValue(@WebParam(name="sVal") String sVal) {
    _cell.setCellValue(sVal);
  }
  /**
   * �u�b�N�t�@�C���o��
   *@param fName �o�̓t�@�C����
   */
  public void write(@WebParam(name="fname") String fName) {
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream(fName);
      _workBook.write(out);
    }catch(IOException e){
      System.out.println(e.toString());
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println(e.toString());
      }
    }
  }
}
