import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
/**
 * JBasic�����^�ɂ��邳���Ȃ����b�p�[�N���X
 */
public class JBasWrapper {
  private Workbook _workBook = null;
  private Sheet _sheet = null;
  private Font _font = null;
  private CellStyle _style = null;
  private Row _row = null;
  private Cell _cell = null;
  private String _mode = "";
  /** �R���X�g���N�^ */
  public JBasWrapper(){
  }
  /** 
   * ���[�N�V�[�g�A�s�A�Z������
   *@param sMode ���[�h
   *@param sName ���[�N�V�[�g�̖��O
   */ 
  public void createWorkSheetAndRowAndCell(String sMode, String sName) {
    _mode = sMode;
    _workBook = _mode.equals("2003") ? new HSSFWorkbook() : new XSSFWorkbook();
    _font = _workBook.createFont();
    _style = _workBook.createCellStyle();
    _sheet = _workBook.createSheet(sName);
    _row = _sheet.createRow(0);
    _cell = _row.createCell(0);
  }
  /**
   * �t�H���g�̎w��ƃX�^�C���ݒ�
   *@param fontName �t�H���g��
   *@param po �����̃|�C���g
   *@param col �����F
   */
  public void setFontAndStyle(String fontName, int po, int col) {
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
  public void setCellValue(String sVal) {
    _cell.setCellValue(sVal);
  }
  /**
   * �u�b�N�t�@�C���o��
   *@param fName �o�̓t�@�C����
   */
  public void write(String fName) {
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream(fName + (_mode.equals("2003") ? ".xls" : ".xlsx"));
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
