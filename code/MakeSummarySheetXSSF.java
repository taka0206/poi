import java.io.*;
import java.util.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.util.*;

/**
 * POI XSSF ->Excel�h�L�������g���색�C�u����
 * �N���X�\���T�v������[�N�V�[�g�V�[�g�쐬
 */
class MakeSummarySheet {
  
  XSSFWorkbook _workBook = null;
  XSSFSheet _sheet = null;
  XSSFPatriarch _patr = null;
  
  /** �R���X�g���N�^�[ */
  public MakeSummarySheet() { 
  }
  /**
   * �t�H���g�����T�u
   *@param point �����̃|�C���g��
   *@param center �Z���^�����O���邩�ǂ����̃t���O
   */
  protected XSSFCellStyle getCellStyle(short point, boolean center) {
    try {
      XSSFCellStyle st = _workBook.createCellStyle();
      if (center == true){
        st.setAlignment(XSSFCellStyle.ALIGN_CENTER);
      }
      XSSFFont fnt = _workBook.createFont();
      fnt.setFontName("�l�r �o�S�V�b�N");
      fnt.setFontHeightInPoints(point);
      st.setFont(fnt);
      return st;
    } catch (Exception e) {
      System.out.println(e.toString());
    }
    return null;

  }
  /**
   * �R�����g�I�u�W�F�N�g�����T�u
   *@param comment �R�����g�ɐݒ肵����������
   */
  protected XSSFComment getComment(String comment) {
    XSSFComment cmt = 
      _patr.createComment(new XSSFClientAnchor(0, 0, 0, 0, 
                (short)1, 1, (short) 8, 6));
    XSSFRichTextString rt = new XSSFRichTextString(comment);
    XSSFFont fnt = _workBook.createFont();
    fnt.setFontName("�l�r �o�S�V�b�N");
    fnt.setFontHeightInPoints((short)14);
    fnt.setColor((short)XSSFColor.BLUE.index);
    fnt.setItalic(true);
    fnt.setBoldweight((short)10);
    rt.applyFont(fnt);
    cmt.setString(rt); 
    cmt.setAuthor(new String("�ۉ� �F�i"));
    return cmt;
  }
  /** �ꊇ�r���`�揈��
   *@param stRow �J�nRow
   *@param edRow �I��Row
   *@param stCell �J�nCell
   *@param edCell �I��Cell
   */
  protected void drawLines(int stRow, int edRow, int stCell, int edCell) {
    // ��r���͊J�n�s�̃Z���̂�
    XSSFRegionUtil.setBorderTop(XSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT ,
        new Region(stRow, (short)stCell, stRow, (short)edCell), _sheet, _workBook);
    // ���r���͏I���s�̃Z���̂�
    XSSFRegionUtil.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT ,
        new Region(edRow, (short)stCell, edRow, (short)edCell), _sheet, _workBook);
    // ���r���͊e�s�̊J�n�Z���̂�
    XSSFRegionUtil.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT,
        new Region(stRow, (short)stCell, edRow, (short)stCell), _sheet, _workBook);
    // �E�r���͊e�s�̏I���Z���̂�
    XSSFRegionUtil.setBorderRight(XSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT,
        new Region(stRow, (short)edCell, edRow, (short)edCell), _sheet, _workBook);
  }
  /** �����̎��s */
  public void Run() {
    // ���[�N�u�b�N�̐���
    _workBook = new XSSFWorkbook();
    // ���[�N�V�[�g�̐���
    _sheet = _workBook.createSheet("POI���C�u�����T�v");
    _patr = _sheet.createDrawingPatriarch();

    // Row�̈ꊇ����
    for (int i=0;i<25;i++) {
      XSSFRow row = _sheet.createRow(i);
      //cell�̈ꊇ����
      for (int j=0;j<20;j++) {
        row.createCell((short)j);
      }
    }
    _sheet.getRow(22).getCell(3).setCellValue("���[�N�u�b�N�S��");
    _sheet.getRow(22).getCell(3).setCellStyle(getCellStyle((short)36,false));
    // ���[�N�u�b�N�R�����g
    XSSFComment cmtBook = getComment("XSSFWorkbook");
    _sheet.getRow(22).getCell(3).setCellComment(cmtBook);

    // �傫���r��������
    drawLines(0,19,0,9);

    // �V�[�g�̕����ƃR�����g�̐ݒ�
    _sheet.getRow(3).getCell(3).setCellValue("���[�N�V�[�g");
    _sheet.getRow(3).getCell(3).setCellStyle(getCellStyle((short)24,false));
    XSSFComment cmtSheet = getComment("XSSFSheet");
    _sheet.getRow(3).getCell(3).setCellComment(cmtSheet);

    // Row�̕����ƌr���`��ƃR�����g�̐ݒ�
    _sheet.addMergedRegion(new Region(8, (short)0, 8, (short)9));
    _sheet.getRow(8).getCell(0).setCellValue("�s(Row)");
    _sheet.getRow(8).getCell(0).setCellStyle(getCellStyle((short)12,true));
    drawLines(8,8,0,9);
    XSSFComment cmtRow = getComment("XSSFRow");
    _sheet.getRow(8).getCell(0).setCellComment(cmtRow);

    // Cell�̕�����ƌr���`��ƃR�����g�̐ݒ�
    _sheet.getRow(16).getCell(1).setCellValue("�Z��(Cell)");
    _sheet.getRow(16).getCell(1).setCellStyle(getCellStyle((short)9,false));
    // �C�ӂ̃Z���Ɍr��������
    drawLines(16,16,1,1);
    // Cell�R�����g
    XSSFComment cmtCell = getComment("XSSFCell");
    _sheet.getRow(16).getCell(1).setCellComment(cmtCell);

    // ���[�W�����̕�����ƌr���`��ƃR�����g�̐ݒ�
    _sheet.addMergedRegion(new Region(16, (short)5, 17, (short)7));
    _sheet.getRow(16).getCell(5).setCellValue("�}�[�W�h���[�W����(MergedRegion)");
    _sheet.getRow(16).getCell(5).setCellStyle(getCellStyle((short)9,true));
    drawLines(16,17,5,7);
    // ���[�W�����R�����g
    XSSFComment cmtRegion = getComment("�}�[�W���ꂽXSSFCell");
    _sheet.getRow(16).getCell(5).setCellComment(cmtRegion);

    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream("./�T�v.xls");
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
    System.out.println("done!");
  }
  /**
   * �G���g���[�|�C���g
   */
  public static void main(String[] args){
    new MakeSummarySheet().Run();
  }
}
