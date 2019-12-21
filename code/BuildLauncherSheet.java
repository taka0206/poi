import java.io.*;
import org.apache.poi.ss.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.*;
/**
 * �T���v�������`���[�V�[�g�쐬
 */
class BuildLauncherSheet {

  protected String _mode;   // ���샂�[�h
  protected Workbook _workBook; // �����`���[���[�N�u�b�N�̃C���X�^���X
  protected int _classPos;    // �N���X���̌��ʒu
  protected String[] _breakKeys; // �L�[�u���[�N���o�ޔ�̈�
  protected int[] _breakPos; // �L�[�u���[�N�s�ԍ��ޔ�̈�
  /** 
   * ��������
   */
  protected boolean init() {
    FileInputStream fis = null;
    // ���[�N�u�b�N��ǂݍ���
    _workBook = null;
    try {
      fis = new FileInputStream( _mode.equals("2003") ? "./input/SampleLauncherORG.xls" : "./input/SampleLauncherORG.xlsm");
      _workBook = _mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + e.toString());
      return false;
    }
    // �N���X���̌��ʒu���擾
    _classPos = (int)_workBook.getSheet("�f�[�^�V�[�g").getRow(0).getCell(1).getNumericCellValue();
    System.out.println("codePos = " + _classPos);
    // �L�[�u���[�N���o�ޔ�̈����������B
    _breakKeys = new String[_classPos];
    // �L�[�u���[�N���o�ޔ�̈揉����
    for (int i=0; i<_classPos; i++) {
      _breakKeys[i] = "";
    }
    // �L�[�u���C�N�s�ԍ��ޔ�̈����������B
    _breakPos = new int[_classPos];
    // �L�[�u���[�N�s�ԍ��ޔ�̈揉����
    for (int i=0; i<_classPos; i++) {
      _breakPos[i] = -1;
    }
    
    return true;
  }
  /** 
   * �����`���[�V�[�g�쐬����
   */
  protected void buildSheet() {
    CreationHelper cHelper = _workBook.getCreationHelper();
    // �[�������N�Z���p�X�^�C�����쐬���Ă���
    CellStyle style = _workBook.createCellStyle();
    Font fnt = _workBook.createFont();
    fnt.setFontName("�l�r �S�V�b�N");
    fnt.setFontHeightInPoints((short)9);
    fnt.setColor((short)org.apache.poi.hssf.util.HSSFColor.BLUE.index);
    fnt.setUnderline(Font.U_SINGLE);
    style.setFont(fnt);
    // ������x���Z���X�^�C�����쐬(�������͍폜)
    CellStyle styleMi = _workBook.createCellStyle();
    Font fntMi = _workBook.createFont();
    fntMi.setFontName("�l�r �S�V�b�N");
    fntMi.setFontHeightInPoints((short)9);
    fntMi.setColor((short)org.apache.poi.hssf.util.HSSFColor.RED.index);
    styleMi.setFont(fntMi);

    // �f�[�^�V�[�g�ƃ����`���[�V�[�g�̎擾
    Sheet dSheet = _workBook.getSheet("�f�[�^�V�[�g");
    Sheet lSheet = _workBook.getSheet("�����`���[�V�[�g");
    // �f�[�^�V�[�g�A�����`���[�V�[�g�Ƃ�3�s�ڂ��珈��
    for (int i=2; i<=dSheet.getLastRowNum(); i++) {
        // �f�[�^�V�[�g����Row�̎擾
        Row dRow = dSheet.getRow(i);
        // �����`���[�V�[�g�ɍs����
        Row lRow = lSheet.createRow(i);
      // �������
      for (int j=0; j<_classPos; j++) {
        Cell cell = dRow.getCell(j);
        if (cell != null) {
          String s = cell.getStringCellValue();
          System.out.println(s);
          if (s.equals(_breakKeys[j]) == false) {
            // ���o���u���[�N����΃����`���[�V�[�g�ɐݒ�
            lRow.createCell(j).setCellValue(s);
            System.out.println("�����`���[�V�[�g�ɍ��ڐݒ�");
            if (_breakPos[j] != -1) {
              if ( (i - _breakPos[j]) > 1 ) {
                // �Ԃ��J���Ă���ꍇCell���c�Ƀ}�[�W����B
                lSheet.addMergedRegion(new CellRangeAddress(_breakPos[j],i-1,j,j));
              }
              
            }
            _breakPos[j] = i; // �L�[�u���[�N�s�ԍ��Ɍ��݂̍s��ݒ�
          }
          _breakKeys[j] = s;
        }
      }
      // �N���X���֘A����
      Cell cell = dRow.getCell(_classPos);
      if (cell != null) {
        String className = cell.getStringCellValue();
        if (className.equals("") == false) {
          if (className.equals("�Ȃ�") == false && className.equals("������") == false) {
            lRow.createCell(_classPos).setCellValue(className);
            boolean bBook1 = dRow.getCell(_classPos + 1).getBooleanCellValue(); // Book�����t���O
            lRow.createCell(_classPos + 1).setCellValue(bBook1);
            boolean both = dRow.getCell(_classPos + 2).getBooleanCellValue(); // 2003/2007���o�[�W��������
            boolean bonly = dRow.getCell(_classPos + 3).getBooleanCellValue();  // �\�[�X�̂݃t���O
            // �\�[�X�t�@�C���Q��Cell�쐬
            Cell fCell = lRow.createCell(_classPos + 2);
            fCell.setCellValue("�\�[�X�t�@�C���Q��");
            fCell.setCellStyle(style);
            if (bonly == false) {
              // �r���hCell�쐬
              Cell bCell = lRow.createCell(_classPos + 3);
              bCell.setCellValue("�r���h");
              bCell.setCellStyle(style);
              // ���s(2003)Cell�쐬
              Cell exCell2003 = lRow.createCell(_classPos + 4);
              exCell2003.setCellValue("���s(2003)");
              exCell2003.setCellStyle(style);
							// ���ʕ\��(2003)�@Book�����t���O��true�̏ꍇ�̂�
							Cell result2003 = lRow.createCell(_classPos + 5);
							if (bBook1) {
	              result2003.setCellValue("���ʕ\��(2003)");
	              result2003.setCellStyle(style);
							}
							else {
	              result2003.setCellValue("�\");
							}
              if (both) {
                // ���s(2007)Cell�쐬
                Cell exCell2007 = lRow.createCell(_classPos + 6);
                exCell2007.setCellValue("���s(2007)");
                exCell2007.setCellStyle(style);
	              Cell result2007 = lRow.createCell(_classPos + 7);
								// ���ʕ\��(2007) Book�����t���O��true�̏ꍇ�̂�
								if (bBook1) {
	                result2007.setCellValue("���ʕ\��(2007)");
	                result2007.setCellStyle(style);
								}
								else {
									result2007.setCellValue("�\");
								}
              }
            }
          }
          else {
            Cell cellNoSource = lRow.createCell(_classPos);
            cellNoSource.setCellValue(className);
            // ������̏ꍇ
            if ( className.equals("������") == true) {
              // ������ԐF��
              cellNoSource.setCellStyle(styleMi);
            }
          }
        }
      }
    }
    // �Ō�̃Z���}�[�W���s���B
    for (int i=0;i<_classPos-1; i++) {
      if (_breakPos[i] != -1 && _breakPos[i] != lSheet.getLastRowNum()) {
        lSheet.addMergedRegion(new CellRangeAddress(_breakPos[i],lSheet.getLastRowNum(),i,i));
      }
    }
    // �N���X���Ȍ�̃J�������������ݒ�ɂ���B
    for (int i=_classPos ; i<_classPos + 6; i++) {
      lSheet.autoSizeColumn(i);
    }
    // Book�o�̓t���O�J�������\����
    lSheet.setColumnHidden(_classPos + 1, true);
    // �V�[�g�𕪊�����B
    lSheet.createFreezePane(_classPos + 2, 2);
    // �����`���[�V�[�g�̍\�z���I���΃f�[�^�V�[�g���폜����B
    _workBook.removeSheetAt(_workBook.getSheetIndex("�f�[�^�V�[�g"));
    // ��Ɨp�V�[�g���\���ɂ���B
    _workBook.setSheetHidden(_workBook.getSheetIndex("��Ɨp�V�[�g"), true);
  }
  /**
   * Excel�u�b�N�o�͏���
   */
  protected void write() {
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( _mode.equals("2003") ? this.getClass().getName() + "_Book1.xls" : 
                      this.getClass().getName() + "_Book1.xlsm");
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
  /** �����̎��s
   * @param ���[�h
   */
  public void Run(String mode) {

    _mode = mode;
    // ��������
    if (init() == false) {
      return;
    }
    // �����`���[�V�[�g�쐬
    buildSheet();
    // ���[�N�u�b�N�����o��
    write();
  

    System.out.println("done!");
  }
  /** �G���g���[�|�C���g */
  public static void main(String[] args) {
    if (args.length != 1) {
      System.out.println("�G���[�F���[�h���w�肵�Ă��������B");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }
    new BuildLauncherSheet().Run(args[0]);
    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
