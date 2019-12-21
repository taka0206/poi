import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * POI HSSF ->Excel�h�L�������g���색�C�u����
 * �N���X�\���T�v������[�N�V�[�g�V�[�g�쐬
 */
class MakeSummarySheet {
  
  HSSFWorkbook _workBook = null;
  HSSFSheet _sheet = null;
  HSSFPatriarch _patr = null;
  
  /** �R���X�g���N�^�[ */
  public MakeSummarySheet() { 
  }
  /**
   * �t�H���g�����T�u
   *@param point �����̃|�C���g��
   *@param center �Z���^�����O���邩�ǂ����̃t���O
   */
  protected HSSFCellStyle getCellStyle(short point, 
                          boolean center) {
    try {
      HSSFCellStyle st = _workBook.createCellStyle();
      if (center == true){
        st.setAlignment(HSSFCellStyle.ALIGN_CENTER);
      }
      HSSFFont fnt = _workBook.createFont();
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
  protected HSSFComment getComment(String comment) {
    HSSFComment cmt = 
      _patr.createComment(new HSSFClientAnchor(
                0, 0, 0, 0, 
                (short)1, 1, (short) 8, 6));
    HSSFRichTextString rt = 
              new HSSFRichTextString(comment);
    HSSFFont fnt = _workBook.createFont();
    fnt.setFontName("�l�r �o�S�V�b�N");
    fnt.setFontHeightInPoints((short)14);
    fnt.setColor((short)HSSFColor.BLUE.index);
    fnt.setItalic(true);
    fnt.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
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
  protected void drawLines(int stRow, int edRow, 
                            int stCell, int edCell) {
    // ��r���͊J�n�s�̃Z���̂�
    HSSFRegionUtil.setBorderTop(
        HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT ,
        new Region(stRow, (short)stCell, 
                    stRow, (short)edCell),
        _sheet, _workBook);
    // ���r���͏I���s�̃Z���̂�
    HSSFRegionUtil.setBorderBottom(
        HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT ,
        new Region(edRow, (short)stCell,
                  edRow, (short)edCell),
        _sheet, _workBook);
    // ���r���͊e�s�̊J�n�Z���̂�
    HSSFRegionUtil.setBorderLeft(
        HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT,
        new Region(stRow, (short)stCell, 
                  edRow, (short)stCell),
        _sheet, _workBook);
    // �E�r���͊e�s�̏I���Z���̂�
    HSSFRegionUtil.setBorderRight(
        HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT,
        new Region(stRow, (short)edCell,
                  edRow, (short)edCell), 
        _sheet, _workBook);
  }
  /** �����̎��s */
  public void Run() {
    // ���[�N�u�b�N�̐���
    _workBook = new HSSFWorkbook();
    // ���[�N�V�[�g�̐���
    _sheet = _workBook.createSheet("POI���C�u�����T�v");
    _patr = _sheet.createDrawingPatriarch();

    // Row�̈ꊇ����
    for (int i=0;i<25;i++) {
      HSSFRow row = _sheet.createRow(i);
      //cell�̈ꊇ����
      for (int j=0;j<20;j++) {
        row.createCell((short)j);
      }
    }
    _sheet.getRow(22).getCell(3).setCellValue(
            "���[�N�u�b�N�S��");
    _sheet.getRow(22).getCell(3).setCellStyle(
            getCellStyle((short)36,false));
    // ���[�N�u�b�N�R�����g
    HSSFComment cmtBook = getComment("HSSFWorkbook");
    _sheet.getRow(22).getCell(3).setCellComment(
              cmtBook);

    // �傫���r��������
    drawLines(0,19,0,9);

    // �V�[�g�̕����ƃR�����g�̐ݒ�
    _sheet.getRow(3).getCell(3).setCellValue(
              "���[�N�V�[�g");
    _sheet.getRow(3).getCell(3).setCellStyle(
                  getCellStyle((short)24,false));
    HSSFComment cmtSheet = getComment("HSSFSheet");
    _sheet.getRow(3).getCell(3).setCellComment(
              cmtSheet);

    // Row�̕����ƌr���`��ƃR�����g�̐ݒ�
    _sheet.addMergedRegion(
          new Region(8, (short)0, 8, (short)9));
    _sheet.getRow(8).getCell(0).setCellValue(
              "�s(Row)");
    _sheet.getRow(8).getCell(0).setCellStyle(
          getCellStyle((short)12,true));
    drawLines(8,8,0,9);
    HSSFComment cmtRow = getComment("HSSFRow");
    _sheet.getRow(8).getCell(0).setCellComment(
          cmtRow);

    // Cell�̕�����ƌr���`��ƃR�����g�̐ݒ�
    _sheet.getRow(16).getCell(1).setCellValue(
          "�Z��(Cell)");
    _sheet.getRow(16).getCell(1).setCellStyle(
                getCellStyle((short)9,false));
    // �C�ӂ̃Z���Ɍr��������
    drawLines(16,16,1,1);
    // Cell�R�����g
    HSSFComment cmtCell = getComment("HSSFCell");
    _sheet.getRow(16).getCell(1).setCellComment(
                cmtCell);

    // ���[�W�����̕�����ƌr���`��ƃR�����g�̐ݒ�
    _sheet.addMergedRegion(
      new Region(16, (short)5, 17, (short)7));
    _sheet.getRow(16).getCell(5).setCellValue(
          "�}�[�W�h���[�W����(MergedRegion)");
    _sheet.getRow(16).getCell(5).setCellStyle(
          getCellStyle((short)9,true));
    drawLines(16,17,5,7);
    // ���[�W�����R�����g
    HSSFComment cmtRegion = 
            getComment("�}�[�W���ꂽHSSFCell");
    _sheet.getRow(16).getCell(5).setCellComment(
        cmtRegion);

    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream(
        this.getClass().getName() + "_Book1.xls");
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

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
