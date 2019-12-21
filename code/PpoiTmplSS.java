import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 

/**
 * �e���v���[�g������Excel���[���쐬����N���X
 */
public class PpoiTmplSS {
  // �e���v���[�g�u�b�N��
  protected String _tmplBookName;
  // �o�̓u�b�N��
  protected String _outBookName;
  // �o�͗p�f�[�^�z��
  protected ArrayList _itemAry;
  // ���[�N�u�b�N�C���^�[�t�F�[�X
  protected Workbook _workBook = null;
  // �ʒu���z��(�n�b�V��)
  protected Hashtable<String,PosInfo> _posTbl = 
        new Hashtable<String,PosInfo>();
  // �֐����z��
  protected ArrayList _funcTbl = new ArrayList();
  // ���샂�[�h "2003" or "2007"
  protected String _mode;
  /**
   * �C���i�[�N���X �ʒu���
   */
  protected class PosInfo {
    public String _itemName;  // ���ږ���
    public int _row;          // �s
    public int _col;          // ��
    public String _type;      // �^�C�v
    public boolean _isArray;  // �z��t���O
    public int _arrayMax;     // �z��ő�l
    public int _incValue;     // �s����
    /** �R���X�g���N�^�[
     *@param itemName ���ږ���
     *@param row      �s
     *@param col      ��
     *@param type     �^�C�v
     *@param isArray  �z��H
     *@param arrayMax �z��ő�l
     *@param incValue �s����
     */
    public PosInfo( String itemName
                    ,int row
                    ,int col
                    ,String type
                    ,boolean isArray
                    ,int arrayMax
                    ,int incValue
                  ) {
      this._itemName = itemName;
      this._row = row;
      this._col = col;
      this._type = type;
      this._isArray = isArray;
      this._arrayMax = arrayMax;
      this._incValue = incValue;
    }
  }
  /**
   * �C���i�[�N���X�@�֐����
   */
  protected class FuncInfo {
    public String _funcName;  // ���ږ���(�g��Ȃ�)
    public int _row;          // �s
    public int _col;          // ��
    public boolean _isArray;  // �z��t���O
    public int _arrayMax;     // �z��ő�l
    public int _incValue;     // �s����
    /**
     * �R���X�g���N�^�[
     *@param funcName ���ږ���
     *@param row      �s
     *@param col      ��
     *@param isArray  �z��H
     *@param arrayMax �z��ő�l
     *@param incValue �s����
     */
    public FuncInfo( String funcName
                    ,int row
                    ,int col
                    ,boolean isArray
                    ,int arrayMax
                    ,int incValue
                  ) {
      this._funcName = funcName;
      this._row = row;
      this._col = col;
      this._isArray = isArray;
      this._arrayMax = arrayMax;
      this._incValue = incValue;
    }
  }
  /**
   * �R���X�g���N�^�[
   *@param mode         ���샂�[�h
   *@param tmplBookName �e���v���[�g�u�b�N�t�@�C����
   *@param outBookName  �o�̓u�b�N�t�@�C����
   *@param itemAry      �o�͗p�f�[�^�z��
   */
  public PpoiTmplSS(String mode, 
                    String tmplBookName,
                    String outBookName,
                    ArrayList itemAry){
    this._mode = mode;
    this._tmplBookName = tmplBookName;
    this._outBookName = outBookName;
    this._itemAry = itemAry;
  }
  /**
   * �e���v���[�g�ǂݍ��ݏ���
   *@return �ǂݍ��݂ɐ��������ꍇ��True 
   */
  protected boolean readTemplate() {
    System.out.println("�e���v���[�g�t�@�C����ǂݍ��݂܂��B");
    // �e���v���[�g�ǂݍ���
    try {
      if (_mode.equals("2003") ) {
        _workBook = new HSSFWorkbook(
            new FileInputStream(_tmplBookName));
      }
      else {
        _workBook = new XSSFWorkbook(
            new FileInputStream(_tmplBookName));
      }
    }
    catch( FileNotFoundException e ) {
      System.out.println(
        "�e���v���[�g�u�b�N�t�@�C�������݂��܂���(" + 
        _tmplBookName + ")�B");
      return false;
    }
    catch( IOException e ) {
      System.out.println(
        "�e���v���[�g�u�b�N�t�@�C���̓ǂݍ��݂Ɏ��s���܂���(" + 
        _tmplBookName + ")�B" + e.toString());
      return false;
    }
    catch( Exception e ) {
      System.out.println(
        "�e���v���[�g�ǂݍ��݂ŃG���[���������܂����B" + 
        e.toString());
      return false;
    }
    return true;
  }
  /** �ʒu���e�[�u���\�z */
  protected boolean buildPosTable() {
    System.out.println("�ʒu�����\�z���܂��B");
    try {
      Sheet sheet = _workBook.getSheetAt(1);
      _posTbl.clear();
      for (int i=1;; i++) {
        // ���ڐ����ϓ�����\��������̂ŁA
        // Null�s���o������܂Ń��[�v
        Row row = sheet.getRow(i);
        if (row == null) break;
        PosInfo info = new PosInfo(
            row.getCell(0).getStringCellValue()
            ,(int)(row.getCell(1).getNumericCellValue())
            ,(int)(row.getCell(2).getNumericCellValue())
            ,row.getCell(3).getStringCellValue()
            ,row.getCell(4).getBooleanCellValue()
            ,(int)(row.getCell(5).getNumericCellValue())
            ,(int)(row.getCell(6).getNumericCellValue())
            );
        _posTbl.put(info._itemName,info);
      }
    }
    catch (Exception e) {
      System.out.println("�ʒu���̍\�z�ŃG���[���������܂����B" + 
                    e.toString());
      return false;
    }
    return true;
  }
  /** 
   * �֐����e�[�u���\�z
   */
  protected boolean buildFuncTable() {
    System.out.println("�֐������\�z���܂��B");
    Sheet sheet = _workBook.getSheetAt(1);
    _funcTbl.clear();
    try{
        // ���ڐ����ϓ�����\��������̂ŁA
        // Null�s���o������܂Ń��[�v
      for (int i=12;;i++) {
        Row row = sheet.getRow(i);
        if (row == null) break;
        FuncInfo info = new FuncInfo(
            row.getCell(0).getStringCellValue()
            ,(int)(row.getCell(1).getNumericCellValue())
            ,(int)(row.getCell(2).getNumericCellValue())
            ,row.getCell(3).getBooleanCellValue()
            ,(int)(row.getCell(4).getNumericCellValue())
            ,(int)(row.getCell(5).getNumericCellValue())
            );
        _funcTbl.add(info);
      }
    }
    catch (Exception e) {
      System.out.println("�֐����̍\�z�ŃG���[���������܂����B" + 
              e.toString());
      return false;
    }
    return true;
  }
  /**
   * Excel���[�̍쐬
   */
  protected boolean makeExcelDocument() {
    System.out.println("Excel���[�̍쐬���s���܂��B");
    Sheet sheet = _workBook.getSheetAt(0);
    try {
      for (int i=0; i<_itemAry.size(); i++) {
        String line = (String)_itemAry.get(i);
        String[] items = line.split("\t");
        PosInfo info = _posTbl.get(items[0]);
        if (info != null) {
          // �P��̂Ƃ��Ɣz��̂Ƃ��ŏ����𕪂���
          if (info._isArray == false) {
            // �P��̂Ƃ�
            Cell cell = 
              sheet.getRow(info._row).getCell(info._col);
            if (setCellValue(
                  cell,info._type,items[1]) == false ) {
              return false;
            }
          }
          else {
            for (int j=1; j<items.length; j++) {
              // �J��Ԃ��ő�l���f�[�^�������ꍇ�͊��Ă�B
              if (j>info._arrayMax) {
                break;
              }
              Cell cell = sheet.getRow(
                info._row + 
                (j*info._incValue) - 1).getCell(info._col);
              if (setCellValue(cell,
                    info._type,items[j]) == false ) {
                return false;
              }
            }
          }
        }
      }
      // ���[�N�V�[�g�̓��e���ďW�v
      if (resetFuncs() == false) return false;
    }
    catch (Exception e) {
      System.out.println("Excel���[�쐬�ŃG���[���������܂����B" + 
                e.toString());
      return false;
    }
    return true;
  }
  /**
   * Cell�ɒl�ݒ�
   *@param cell �l��ݒ肷��Z��
   *@param type �^�C�v
   *@param value �ݒ肷��l
   */
  protected boolean setCellValue(Cell cell, 
                        String type, String value) {
    try {
      if (type.equals("string")) {
        cell.setCellValue(value);
      }
      else if(type.equals("nummber")) {
        cell.setCellValue(Double.parseDouble(value));
      }
      else if(type.equals("date")) {
        cell.setCellValue(
          DateFormat.getDateInstance().parse(value));
      }
    }
    catch (Exception e) {
      System.out.println("Cell�ւ̒l�ݒ�ŃG���[���������܂����B" +
                e.toString());
      return false;
    }
    return true;
  }
  /**
   * ���ߍ��݊֐��̍Đݒ菈��
   * 2007���[�h�̏ꍇ�̂ݕK�v
   */
  protected boolean resetFuncs() {
    System.out.println("���ߍ��݊֐��̍Đݒ���s���܂��B");
    try {
      for(int i=0; i<_funcTbl.size();i++) {
        FuncInfo tbl = (FuncInfo)_funcTbl.get(i);
        if (tbl._isArray == false) {
          Cell cell = 
            _workBook.getSheetAt(0).getRow(
                tbl._row).getCell(tbl._col);
          String func = cell.getCellFormula();
          cell.setCellFormula(func);
        }
        else {
          int rpos = tbl._row;
          for (int j=0;j<tbl._arrayMax;j++) {
            Cell cell = 
              _workBook.getSheetAt(0).getRow(rpos).getCell(tbl._col);
            String func = cell.getCellFormula();
            cell.setCellFormula(func);
            rpos += tbl._incValue;
          }
        }
      }
    }
    catch(Exception e) {
      System.out.println("�g�ݍ��݊֐��̍Đݒ�ŃG���[���������܂����B" + 
                  e.toString());
    }
    return true;
  }
  /**
   * ExcelBook�����o��
   */
  protected boolean write() {
    System.out.println("�u�b�N�̏����o�����s���܂��B");
    FileOutputStream out = null;
    try{
      out = new FileOutputStream(_outBookName);
      _workBook.removeSheetAt(1);
      _workBook.write(out);
    }catch(IOException e){
      System.out.println(e.toString());
      return false;
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println(e.toString());
        return false;
      }
    }
    return true;
  }
  /**
   * �����̃R���g���[��
   */
  public boolean Run() {
    // �e���v���[�g�u�b�N�ǂݍ���
    if (readTemplate() == false) {
      return false;
    }
    // �ʒu���e�[�u���\�z
    if (buildPosTable() == false) {
      return false;
    }
    // �֐����e�[�u���\�z 2007���[�h�̂Ƃ��̂�
    //if (_mode.equals("2007")){
      if (buildFuncTable() == false) {
        return false;
      }
    //}
    // Excel���[�̍쐬
    if (makeExcelDocument() == false) {
      return false;
    }
    // ExcelBook�����o��
    if (write() == false) {
      return false;
    }
    return true;
  }
  /**
  * �e�X�g�p���[�`��
  *@param args [0]->���샂�[�h�A
  *            [1]->�f�[�^�t�@�C���A
  *            [2]->�e���v���[�g�u�b�N�A
  *            [3]->�o�̓u�b�N
  */
  public static void main(String[] args){

    if (args.length != 1) {
      System.out.println("�p�����[�^�[�G���[�ł��B");
    }
    else if (!args[0].equals("2003") && 
             !args[0].equals("2007")) {
      System.out.println(
        "���샂�[�h��2003�܂���2007���w�肵�Ă��������B");
    }
    else {
      try {
        ArrayList itemArray = new ArrayList();
        itemArray.clear();
        FileReader fr = new FileReader("./input/PpoiTmpl.dat");
        BufferedReader br = new BufferedReader(fr);
        String line;
        while ((line = br.readLine()) != null) {
          itemArray.add(line);
        }
        br.close();
        fr.close();
        if (new PpoiTmplSS(args[0],
                args[0].equals("2003") ? 
                "./input/nohinsyo_tmpl.xls" :
                "./input/nohinsyo_tmpl.xlsx",
                args[0].equals("2003") ? 
                "./PpoiTmplSS_Book1.xls" :
                "./PpoiTmplSS_Book1.xlsx",
                itemArray).Run()==true) {
          System.out.println("����I��");
        }
        else {
          System.out.println("�ُ�I��");
        }
      }
      catch (Exception e) {
        System.out.println("�G���[���������܂����B" + 
                e.toString());
      } 
    }
  }
}
