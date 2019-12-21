import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

/**
 * POI HSSF/XSSF���� ->Excel�h�L�������g���색�C�u����
 * �N���X�\���T�v������[�N�V�[�g�V�[�g�쐬
 */
class MakeSummarySheetSS {
  
  protected String _mode;
  // ���[�N�u�b�N�C���^�[�t�F�[�X
  protected Workbook _workBook = null;
  // �V�[�g�C���^�[�t�F�[�X
  protected Sheet _sheet = null;
  // Patriarch�I�u�W�F�N�g 2003�̏ꍇ�̂�
  protected HSSFPatriarch _patr2003 = null;
  // Drawing�I�u�W�F�N�g 2007�̏ꍇ�̂�
  protected XSSFDrawing _patr2007 = null;

  /** 
   * �R���X�g���N�^�[
   *@param mode ���샂�[�h
  */
  public MakeSummarySheetSS(String mode) { 
    _mode = mode;
  }
  /**
   * �t�H���g�����T�u
   *@param point �����̃|�C���g��
   *@param center �Z���^�����O���邩�ǂ����̃t���O
   */
  protected void setUserCellStyle(Cell cel, 
                short point, boolean center) {
    try {
      CellStyle st = cel.getCellStyle();
      if (center == true){
        st.setAlignment(CellStyle.ALIGN_CENTER);
      }
      Font fnt = _workBook.createFont();
      fnt.setFontName("�l�r �o�S�V�b�N");
      fnt.setFontHeightInPoints(point);
      st.setFont(fnt);
    } catch (Exception e) {
      System.out.println(e.toString());
    }
  }
  /**
   * �R�����g�I�u�W�F�N�g�����T�u
   *@param comment �R�����g�ɐݒ肵����������
   */
  protected Comment getComment(String comment) {

    Comment cmt = null;
    RichTextString rt = null; 

    if (_mode.equals("2003")) {
      cmt = _patr2003.createComment(
              new HSSFClientAnchor(0, 0, 0, 0, 
                (short)1, 1, (short) 8, 6));
      rt = new HSSFRichTextString(comment);
    }
    else {
      cmt = _patr2007.createCellComment(
              new XSSFClientAnchor(0, 0, 0, 0,
                (short)1, 1, (short) 8, 6));
      rt = new XSSFRichTextString(comment);
    }
    Font fnt = _workBook.createFont();
    fnt.setFontName("�l�r �o�S�V�b�N");
    fnt.setFontHeightInPoints((short)14);
    fnt.setColor((short)HSSFColor.BLUE.index);
    fnt.setItalic(true);
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
    for(int i=stCell; i<=edCell;i++) {
      Cell cel = _sheet.getRow(stRow).getCell(i);
      CellStyle styleU = cel.getCellStyle();
      styleU.setBorderTop(
        HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT);
      cel.setCellStyle(styleU);
    }
    // ���r���͏I���s�̃Z���̂�
    for(int i=stCell; i<=edCell;i++) {
      Cell cel = _sheet.getRow(edRow).getCell(i);
      CellStyle styleD = cel.getCellStyle();
      styleD.setBorderBottom(
        HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT);
      cel.setCellStyle(styleD);
    }
    // ���r���͊e�s�̊J�n�Z���̂�
    for(int i=stRow; i<=edRow;i++) {
      Cell cel = _sheet.getRow(i).getCell(stCell);
      CellStyle styleL = cel.getCellStyle();
      styleL.setBorderLeft(
        HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT);
      cel.setCellStyle(styleL);
    }
    // �E�r���͊e�s�̏I���Z���̂�
    for(int i=stRow; i<=edRow;i++) {
      Cell cel = _sheet.getRow(i).getCell(edCell);
      CellStyle styleR = cel.getCellStyle();
      styleR.setBorderRight(
        HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT);
      cel.setCellStyle(styleR);
    }
  }
  /** �����̎��s */
  public void Run() {
    Cell wCell = null;
    // ���[�N�u�b�N�̐���
    if (_mode.equals("2003")) {
      _workBook = new HSSFWorkbook();
    }
    else if (_mode.equals("2007")) {
      _workBook = new XSSFWorkbook();
    }
    else {
      System.out.println(
            "���[�h��2003��2007���w�肵�܂��B");
      return;
    }
    // ���[�N�V�[�g�̐���
    _sheet = _workBook.createSheet("POI���C�u�����T�v");
    if (_mode.equals("2003")) {
      _patr2003 = 
        ((HSSFSheet)_sheet).createDrawingPatriarch();
    }
    else {
      _patr2007 = 
        ((XSSFSheet)_sheet).createDrawingPatriarch();
    }
    // Row�̈ꊇ����
    for (int i=0;i<25;i++) {
      Row row = _sheet.createRow(i);
      //cell�̈ꊇ����
      //style���쐬���Ă���
      for (int j=0;j<20;j++) {
        Cell cel = row.createCell((short)j);
        cel.setCellStyle(_workBook.createCellStyle());
      }
    }
    wCell = _sheet.getRow(22).getCell(3);
    wCell.setCellValue("���[�N�u�b�N�S��");
    setUserCellStyle(wCell,(short)36,false);
    // ���[�N�u�b�N�R�����g
    Comment cmtBook = getComment("Workbook");
    wCell.setCellComment(cmtBook);

    // �傫���r��������
    drawLines(0,19,0,9);

    // �r��������
    drawLines(8,8,0,9);

    // �V�[�g�̕����ƃR�����g�̐ݒ�
    wCell = _sheet.getRow(3).getCell(3);
    wCell.setCellValue("���[�N�V�[�g");
    setUserCellStyle(wCell,(short)24,false);
    Comment cmtSheet = getComment("Sheet");
    wCell.setCellComment(cmtSheet);

    // Row�̕����ƌr���`��ƃR�����g�̐ݒ�
    _sheet.addMergedRegion(
        new org.apache.poi.ss.util.CellRangeAddress(
              8, 8, 0, 9));
    wCell = _sheet.getRow(8).getCell(0);
    wCell.setCellValue("�s(Row)");
    setUserCellStyle(wCell,(short)12,true);
    Comment cmtRow = getComment("Row");
    wCell.setCellComment(cmtRow);

    // �r��������
    drawLines(16,16,1,1);

    // Cell�̕�����ƌr���`��ƃR�����g�̐ݒ�
    wCell = _sheet.getRow(16).getCell(1);
    wCell.setCellValue("�Z��(Cell)");
    setUserCellStyle(wCell,(short)9,false);
    // Cell�R�����g
    Comment cmtCell = getComment("Cell");
    wCell.setCellComment(cmtCell);

    // �r��������
    drawLines(16,17,5,7);

    // ���[�W�����̕�����ƌr���`��ƃR�����g�̐ݒ�
    _sheet.addMergedRegion(
        new org.apache.poi.ss.util.CellRangeAddress(
              16, 17, 5, 7));
    wCell = _sheet.getRow(16).getCell(5);
    wCell.setCellValue("�}�[�W�h���[�W����(MergedRegion)");
    setUserCellStyle(wCell,(short)9,true);
    // ���[�W�����R�����g
    Comment cmtRegion = getComment("�}�[�W���ꂽCell");
    wCell.setCellComment(cmtRegion);

    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      if (_mode.equals("2003")) {
        out = new FileOutputStream(
          this.getClass().getName() + "_Book1.xls");
      }
      else {
        out = new FileOutputStream(
          this.getClass().getName() + "_Book1.xlsx");
      }
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
    if (args.length != 1) {
      System.out.println("�G���[�F���[�h���w�肵�ĉ������B");
      return;
    }
    else if ( !args[0].equals("2003") && 
              !args[0].equals("2007") ) {
      System.out.println(
        "�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }
    new MakeSummarySheetSS(args[0]).Run();

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
