import java.io.*;
import org.apache.poi.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * �V�[�g�ɐ}�`��\��t����(hssf�̂�)
 */ 
public class SetShape {

  // Patriarch�I�u�W�F�N�g �V�F�C�v�p
  protected HSSFPatriarch _patr2003 = null;
  // Patriarch�I�u�W�F�N�g �|���S���p
  protected HSSFPatriarch _patr2003P = null;
  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // ���[�N�V�[�g����
    Sheet sheet = workBook.createSheet("�V�F�C�v");
    // �e��V�F�C�v�����
    if (mode.equals("2003")) {
      _patr2003 = ((HSSFSheet)sheet).createDrawingPatriarch();
      // �l�p�`-1
      HSSFClientAnchor anchorRect1 = 
                  new HSSFClientAnchor(0, 0, 0, 0, 
                      (short)1, 1, (short)3, 9);
      // Cell�ɕ����Ĉړ��E���T�C�Y
      anchorRect1.setAnchorType(0); 
      HSSFSimpleShape rShape1 = 
          _patr2003.createSimpleShape(anchorRect1);
      rShape1.setShapeType(
          HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
      // ���̐F��ɂ���B
      rShape1.setLineStyleColor(0x00, 0x00, 0xff);
      // �g����������Ƒ���(2pt)��
      rShape1.setLineWidth(
          HSSFSimpleShape.LINEWIDTH_ONE_PT * 2);
      // �l�p�̒����V�A���ɓh��ׂ�
      rShape1.setFillColor(0x00, 0xff, 0xff);
      // �l�p�`-2
      HSSFClientAnchor anchorRect2 = 
                  new HSSFClientAnchor(0, 0, 0, 0, 
                      (short)4, 1, (short)6, 9);
      // Cell�ɕ����Ĉړ��E���T�C�Y
      anchorRect2.setAnchorType(0); 
      HSSFSimpleShape rShape2 = 
          _patr2003.createSimpleShape(anchorRect2);
      rShape2.setShapeType(
          HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
      // �g����������Ƒ���(2pt)��
      rShape2.setLineWidth(
          HSSFSimpleShape.LINEWIDTH_ONE_PT * 2);
      // �h��ׂ��Ȃ��ɐݒ�
      rShape2.setNoFill(true);
      sheet.createRow(4).createCell(4).setCellValue(
          "�����Ă܂�");
      // �ȉ~�̕`��
      HSSFClientAnchor anchorOval = 
                new HSSFClientAnchor(0, 0, 0, 0, 
                (short)7, 1, (short)10, 9);
      // Cell�ɕ����Ĉړ��E���T�C�Y
      anchorOval.setAnchorType(0);
      HSSFSimpleShape ovalShape = 
        _patr2003.createSimpleShape(anchorOval);
      ovalShape.setShapeType(
          HSSFSimpleShape.OBJECT_TYPE_OVAL);
      // ���̐F��Ԃɂ���B
      ovalShape.setLineStyleColor(0xff, 0x00, 0x00);
      // �g���𑾂�(5pt)��
      ovalShape.setLineWidth(
          HSSFSimpleShape.LINEWIDTH_ONE_PT * 5);
      // �l�p�̒����I�����W�ɓh��ׂ�
      ovalShape.setFillColor(0xff, 0xa5, 0x00);
      // ������`��
      // �`�󖼃e�[�u���̒�`
      String lineNames[] = {
        "����"
       ,"�j��"
       ,"�_��"
       ,"��_����"
       ,"��_����"
       ,"�e���_��"
       ,"�e���j��"
       ,"�e����_����"
       ,"�����̒�����_����"
       ,"�����̒�����_����"
      };
      // �`��e�[�u���̒�`
      int lineStyles[] = {  
        HSSFShape.LINESTYLE_SOLID
       ,HSSFShape.LINESTYLE_DASHSYS
       ,HSSFShape.LINESTYLE_DOTSYS
       ,HSSFShape.LINESTYLE_DASHDOTSYS
       ,HSSFShape.LINESTYLE_DASHDOTDOTSYS
       ,HSSFShape.LINESTYLE_DOTGEL
       ,HSSFShape.LINESTYLE_LONGDASHGEL
       ,HSSFShape.LINESTYLE_DASHDOTGEL
       ,HSSFShape.LINESTYLE_LONGDASHDOTGEL
       ,HSSFShape.LINESTYLE_LONGDASHDOTDOTGEL
      };
      int line = 11;
      for (int i=0; i<10; i++) {
        // �����̏ꍇ�́ACell�̐^�񒆂�����ɂ���悤�Ƀ}�[�W�������B
        HSSFClientAnchor anchorLine = 
              new HSSFClientAnchor(0, 128, 0, 128, 
                (short)1, line, (short)4, line);
        // Cell�ɕ����Ĉړ��E���T�C�Y
        anchorLine.setAnchorType(0); 
        // SimpleShape�̐���
        HSSFSimpleShape lShape = 
          _patr2003.createSimpleShape(anchorLine);
        // �������w��
        lShape.setShapeType(
          HSSFSimpleShape.OBJECT_TYPE_LINE);
        // �����̌`����w��
        lShape.setLineStyle(lineStyles[i]);
        // ���̌`�������
        Cell cell = sheet.createRow(line).createCell(4);
        cell.setCellValue(lineNames[i]);
        line++;
      }
    }
    else {
      // �����ł͏������Ȃ��B
    }
    // �|���S����`��
    // �V�[�g2���쐬
    Sheet sheet2 = workBook.createSheet("�|���S��");
    if (mode.equals("2003")) {
      _patr2003P = 
        ((HSSFSheet)sheet2).createDrawingPatriarch();
      // ClientAnchor�̐���
      HSSFClientAnchor anchorPol = 
          new HSSFClientAnchor(0, 0, 0, 0, 
            (short)1, 1, (short)6, 9);
      // Cell�ɕ����Ĉړ��E���T�C�Y
      anchorPol.setAnchorType(0); 
      // Polygon�C���X�^���X�𐶐�
      HSSFPolygon pol = 
        _patr2003P.createPolygon(anchorPol);
      // �`��̈�w��
      pol.setPolygonDrawArea(100, 100);
      // �e�_��X�AY���W��ݒ肷��B
      pol.setPoints(
        // x���W�̔z��
        new int[]{10,20,30,40,50,60,70,80,90},
        // y���W�̔z��
        new int[]{10,20,30,20,80,10,50,90,40}); 
      // ���̐F���}���[���ɂ���
      pol.setLineStyleColor(0x80, 0x00, 0x00);
      // �~�f�B�A���p�[�v���œh��ׂ�
      pol.setFillColor(0x93,0x70, 0xdb);
    }
    else {
      // �����ł͏������Ȃ��B
    }

    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? this.getClass().getName() + "_Book1.xls" : 
                      this.getClass().getName() + "_Book1.xlsx");
      workBook.write(out);
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
  /** �G���g���[�|�C���g */

  public static void main(String[] args) {

    if (args.length != 1) {
      System.out.println("�G���[�F���[�h���w�肵�ĉ������B");
      return;
    }
    else if ( !args[0].equals("2003") ) {
      System.out.println("�G���[�F���[�h��2003�̂ݎw�肵�ĉ������B");
      return;
    }
    // �����̎��s
    new SetShape().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
