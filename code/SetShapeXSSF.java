import java.io.*;
import org.apache.poi.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * シートに図形を貼り付ける(xssf)
 */ 
public class SetShapeXSSF {

  
  protected class ShapeTypesInfo {
    public String _typeName;
    public int _typeNo;
    public ShapeTypesInfo(String tname, int tno) {
      _typeName = tname;
      _typeNo = tno;
    }
  }

  protected ShapeTypesInfo[] sti = { new ShapeTypesInfo("LINE", ShapeTypes.LINE)
                                    ,new ShapeTypesInfo("LINE_INV", ShapeTypes.LINE_INV)
                                    ,new ShapeTypesInfo("TRIANGLE",ShapeTypes.TRIANGLE) 
                                    ,new ShapeTypesInfo("RT_TRIANGLE", ShapeTypes.RT_TRIANGLE) 
                                    ,new ShapeTypesInfo("RECT", ShapeTypes.RECT) 
                                    ,new ShapeTypesInfo("DIAMOND", ShapeTypes.DIAMOND) 
                                    ,new ShapeTypesInfo("PARALLELOGRAM", ShapeTypes.PARALLELOGRAM) 
                                    ,new ShapeTypesInfo("TRAPEZOID", ShapeTypes.TRAPEZOID) 
                                    ,new ShapeTypesInfo("NON_ISOSCELES_TRAPEZOID", ShapeTypes.NON_ISOSCELES_TRAPEZOID) 
                                    ,new ShapeTypesInfo("PENTAGON", ShapeTypes.PENTAGON) 
                                    ,new ShapeTypesInfo("HEXAGON", ShapeTypes.HEXAGON) 
                                    ,new ShapeTypesInfo("HEPTAGON", ShapeTypes.HEPTAGON) 
                                    ,new ShapeTypesInfo("OCTAGON", ShapeTypes.OCTAGON) 
                                    ,new ShapeTypesInfo("DECAGON", ShapeTypes.DECAGON) 
                                    ,new ShapeTypesInfo("DODECAGON", ShapeTypes.DODECAGON) 
                                    ,new ShapeTypesInfo("STAR_4", ShapeTypes.STAR_4) 
                                    ,new ShapeTypesInfo("STAR_5", ShapeTypes.STAR_5) 
                                    ,new ShapeTypesInfo("STAR_6", ShapeTypes.STAR_6) 
                                    ,new ShapeTypesInfo("STAR_7", ShapeTypes.STAR_7) 
                                    ,new ShapeTypesInfo("STAR_8", ShapeTypes.STAR_8) 
                                    ,new ShapeTypesInfo("STAR_10", ShapeTypes.STAR_10) 
                                    ,new ShapeTypesInfo("STAR_12", ShapeTypes.STAR_12) 
                                    ,new ShapeTypesInfo("STAR_16", ShapeTypes.STAR_16) 
                                    ,new ShapeTypesInfo("STAR_24", ShapeTypes.STAR_24) 
                                    ,new ShapeTypesInfo("STAR_32", ShapeTypes.STAR_32) 
                                    ,new ShapeTypesInfo("ROUND_RECT", ShapeTypes.ROUND_RECT) 
                                    ,new ShapeTypesInfo("ROUND_1_RECT", ShapeTypes.ROUND_1_RECT) 
                                    ,new ShapeTypesInfo("ROUND_2_SAME_RECT", ShapeTypes.ROUND_2_SAME_RECT) 
                                    ,new ShapeTypesInfo("ROUND_2_DIAG_RECT", ShapeTypes.ROUND_2_DIAG_RECT) 
                                    ,new ShapeTypesInfo("SNIP_ROUND_RECT", ShapeTypes.SNIP_ROUND_RECT) 
                                    ,new ShapeTypesInfo("SNIP_1_RECT", ShapeTypes.SNIP_1_RECT) 
                                    ,new ShapeTypesInfo("SNIP_2_SAME_RECT", ShapeTypes.SNIP_2_SAME_RECT) 
                                    ,new ShapeTypesInfo("SNIP_2_DIAG_RECT", ShapeTypes.SNIP_2_DIAG_RECT) 
                                    ,new ShapeTypesInfo("PLAQUE", ShapeTypes.PLAQUE) 
                                    ,new ShapeTypesInfo("ELLIPSE", ShapeTypes.ELLIPSE) 
                                    ,new ShapeTypesInfo("TEARDROP", ShapeTypes.TEARDROP) 
                                    ,new ShapeTypesInfo("HOME_PLATE", ShapeTypes.HOME_PLATE) 
                                    ,new ShapeTypesInfo("CHEVRON", ShapeTypes.CHEVRON) 
                                    ,new ShapeTypesInfo("PIE_WEDGE", ShapeTypes.PIE_WEDGE) 
                                    ,new ShapeTypesInfo("PIE", ShapeTypes.PIE) 
                                    ,new ShapeTypesInfo("BLOCK_ARC", ShapeTypes.BLOCK_ARC) 
                                    ,new ShapeTypesInfo("DONUT", ShapeTypes.DONUT) 
                                    ,new ShapeTypesInfo("NO_SMOKING", ShapeTypes.NO_SMOKING) 
                                    ,new ShapeTypesInfo("RIGHT_ARROW", ShapeTypes.RIGHT_ARROW) 
                                    ,new ShapeTypesInfo("LEFT_ARROW", ShapeTypes.LEFT_ARROW) 
                                    ,new ShapeTypesInfo("UP_ARROW", ShapeTypes.UP_ARROW) 
                                    ,new ShapeTypesInfo("DOWN_ARROW", ShapeTypes.DOWN_ARROW) 
                                    ,new ShapeTypesInfo("STRIPED_RIGHT_ARROW", ShapeTypes.STRIPED_RIGHT_ARROW) 
                                    ,new ShapeTypesInfo("NOTCHED_RIGHT_ARROW", ShapeTypes.NOTCHED_RIGHT_ARROW) 
                                    ,new ShapeTypesInfo("BENT_UP_ARROW", ShapeTypes.BENT_UP_ARROW) 
                                    ,new ShapeTypesInfo("LEFT_RIGHT_ARROW", ShapeTypes.LEFT_RIGHT_ARROW) 
                                    ,new ShapeTypesInfo("UP_DOWN_ARROW", ShapeTypes.UP_DOWN_ARROW) 
                                    ,new ShapeTypesInfo("LEFT_UP_ARROW", ShapeTypes.LEFT_UP_ARROW) 
                                    ,new ShapeTypesInfo("LEFT_RIGHT_UP_ARROW", ShapeTypes.LEFT_RIGHT_UP_ARROW) 
                                    ,new ShapeTypesInfo("QUAD_ARROW", ShapeTypes.QUAD_ARROW) 
                                    ,new ShapeTypesInfo("LEFT_ARROW_CALLOUT", ShapeTypes.LEFT_ARROW_CALLOUT) 
                                    ,new ShapeTypesInfo("RIGHT_ARROW_CALLOUT", ShapeTypes.RIGHT_ARROW_CALLOUT) 
                                    ,new ShapeTypesInfo("UP_ARROW_CALLOUT", ShapeTypes.UP_ARROW_CALLOUT) 
                                    ,new ShapeTypesInfo("DOWN_ARROW_CALLOUT", ShapeTypes.DOWN_ARROW_CALLOUT) 
                                    ,new ShapeTypesInfo("LEFT_RIGHT_ARROW_CALLOUT", ShapeTypes.LEFT_RIGHT_ARROW_CALLOUT) 
                                    ,new ShapeTypesInfo("UP_DOWN_ARROW_CALLOUT", ShapeTypes.UP_DOWN_ARROW_CALLOUT) 
                                    ,new ShapeTypesInfo("QUAD_ARROW_CALLOUT", ShapeTypes.QUAD_ARROW_CALLOUT) 
                                    ,new ShapeTypesInfo("BENT_ARROW", ShapeTypes.BENT_ARROW) 
                                    ,new ShapeTypesInfo("UTURN_ARROW", ShapeTypes.UTURN_ARROW) 
                                    ,new ShapeTypesInfo("CIRCULAR_ARROW", ShapeTypes.CIRCULAR_ARROW) 
                                    ,new ShapeTypesInfo("LEFT_CIRCULAR_ARROW", ShapeTypes.LEFT_CIRCULAR_ARROW) 
                                    ,new ShapeTypesInfo("LEFT_RIGHT_CIRCULAR_ARROW", ShapeTypes.LEFT_RIGHT_CIRCULAR_ARROW) 
                                    ,new ShapeTypesInfo("CURVED_RIGHT_ARROW", ShapeTypes.CURVED_RIGHT_ARROW) 
                                    ,new ShapeTypesInfo("CURVED_LEFT_ARROW", ShapeTypes.CURVED_LEFT_ARROW) 
                                    ,new ShapeTypesInfo("CURVED_UP_ARROW", ShapeTypes.CURVED_UP_ARROW) 
                                    ,new ShapeTypesInfo("CURVED_DOWN_ARROW", ShapeTypes.CURVED_DOWN_ARROW) 
                                    ,new ShapeTypesInfo("SWOOSH_ARROW", ShapeTypes.SWOOSH_ARROW) 
                                    ,new ShapeTypesInfo("CUBE", ShapeTypes.CUBE) 
                                    ,new ShapeTypesInfo("CAN", ShapeTypes.CAN) 
                                    ,new ShapeTypesInfo("LIGHTNING_BOLT", ShapeTypes.LIGHTNING_BOLT) 
                                    ,new ShapeTypesInfo("HEART", ShapeTypes.HEART) 
                                    ,new ShapeTypesInfo("SUN", ShapeTypes.SUN) 
                                    ,new ShapeTypesInfo("MOON", ShapeTypes.MOON) 
                                    ,new ShapeTypesInfo("SMILEY_FACE", ShapeTypes.SMILEY_FACE) 
                                    ,new ShapeTypesInfo("IRREGULAR_SEAL_1", ShapeTypes.IRREGULAR_SEAL_1) 
                                    ,new ShapeTypesInfo("IRREGULAR_SEAL_2", ShapeTypes.IRREGULAR_SEAL_2) 
                                    ,new ShapeTypesInfo("FOLDED_CORNER", ShapeTypes.FOLDED_CORNER) 
                                    ,new ShapeTypesInfo("BEVEL", ShapeTypes.BEVEL) 
                                    ,new ShapeTypesInfo("FRAME", ShapeTypes.FRAME) 
                                    ,new ShapeTypesInfo("HALF_FRAME", ShapeTypes.HALF_FRAME) 
                                    ,new ShapeTypesInfo("CORNER", ShapeTypes.CORNER) 
                                    ,new ShapeTypesInfo("DIAG_STRIPE", ShapeTypes.DIAG_STRIPE) 
                                    ,new ShapeTypesInfo("CHORD", ShapeTypes.CHORD) 
                                    ,new ShapeTypesInfo("ARC", ShapeTypes.ARC) 
                                    ,new ShapeTypesInfo("LEFT_BRACKET", ShapeTypes.LEFT_BRACKET) 
                                    ,new ShapeTypesInfo("RIGHT_BRACKET", ShapeTypes.RIGHT_BRACKET) 
                                    ,new ShapeTypesInfo("LEFT_BRACE", ShapeTypes.LEFT_BRACE) 
                                    ,new ShapeTypesInfo("RIGHT_BRACE", ShapeTypes.RIGHT_BRACE) 
                                    ,new ShapeTypesInfo("BRACKET_PAIR", ShapeTypes.BRACKET_PAIR) 
                                    ,new ShapeTypesInfo("BRACE_PAIR", ShapeTypes.BRACE_PAIR) 
                                    ,new ShapeTypesInfo("STRAIGHT_CONNECTOR_1", ShapeTypes.STRAIGHT_CONNECTOR_1) 
                                    ,new ShapeTypesInfo("BENT_CONNECTOR_2", ShapeTypes.BENT_CONNECTOR_2) 
                                    ,new ShapeTypesInfo("BENT_CONNECTOR_3", ShapeTypes.BENT_CONNECTOR_3) 
                                    ,new ShapeTypesInfo("BENT_CONNECTOR_4", ShapeTypes.BENT_CONNECTOR_4) 
                                    ,new ShapeTypesInfo("BENT_CONNECTOR_5", ShapeTypes.BENT_CONNECTOR_5) 
                                    ,new ShapeTypesInfo("CURVED_CONNECTOR_2", ShapeTypes.CURVED_CONNECTOR_2) 
                                    ,new ShapeTypesInfo("CURVED_CONNECTOR_3", ShapeTypes.CURVED_CONNECTOR_3) 
                                    ,new ShapeTypesInfo("CURVED_CONNECTOR_4", ShapeTypes.CURVED_CONNECTOR_4) 
                                    ,new ShapeTypesInfo("CURVED_CONNECTOR_5", ShapeTypes.CURVED_CONNECTOR_5) 
                                    ,new ShapeTypesInfo("CALLOUT_1", ShapeTypes.CALLOUT_1) 
                                    ,new ShapeTypesInfo("CALLOUT_2", ShapeTypes.CALLOUT_2) 
                                    ,new ShapeTypesInfo("CALLOUT_3", ShapeTypes.CALLOUT_3) 
                                    ,new ShapeTypesInfo("ACCENT_CALLOUT_1", ShapeTypes.ACCENT_CALLOUT_1) 
                                    ,new ShapeTypesInfo("ACCENT_CALLOUT_2", ShapeTypes.ACCENT_CALLOUT_2) 
                                    ,new ShapeTypesInfo("ACCENT_CALLOUT_3", ShapeTypes.ACCENT_CALLOUT_3) 
                                    ,new ShapeTypesInfo("BORDER_CALLOUT_1", ShapeTypes.BORDER_CALLOUT_1) 
                                    ,new ShapeTypesInfo("BORDER_CALLOUT_2", ShapeTypes.BORDER_CALLOUT_2) 
                                    ,new ShapeTypesInfo("BORDER_CALLOUT_3", ShapeTypes.BORDER_CALLOUT_3) 
                                    ,new ShapeTypesInfo("ACCENT_BORDER_CALLOUT_1", ShapeTypes.ACCENT_BORDER_CALLOUT_1) 
                                    ,new ShapeTypesInfo("ACCENT_BORDER_CALLOUT_2", ShapeTypes.ACCENT_BORDER_CALLOUT_2) 
                                    ,new ShapeTypesInfo("ACCENT_BORDER_CALLOUT_3", ShapeTypes.ACCENT_BORDER_CALLOUT_3) 
                                    ,new ShapeTypesInfo("WEDGE_RECT_CALLOUT", ShapeTypes.WEDGE_RECT_CALLOUT) 
                                    ,new ShapeTypesInfo("WEDGE_ROUND_RECT_CALLOUT", ShapeTypes.WEDGE_ROUND_RECT_CALLOUT) 
                                    ,new ShapeTypesInfo("WEDGE_ELLIPSE_CALLOUT", ShapeTypes.WEDGE_ELLIPSE_CALLOUT) 
                                    ,new ShapeTypesInfo("CLOUD_CALLOUT", ShapeTypes.CLOUD_CALLOUT) 
                                    ,new ShapeTypesInfo("CLOUD", ShapeTypes.CLOUD) 
                                    ,new ShapeTypesInfo("RIBBON", ShapeTypes.RIBBON) 
                                    ,new ShapeTypesInfo("RIBBON_2", ShapeTypes.RIBBON_2) 
                                    ,new ShapeTypesInfo("ELLIPSE_RIBBON", ShapeTypes.ELLIPSE_RIBBON) 
                                    ,new ShapeTypesInfo("ELLIPSE_RIBBON_2", ShapeTypes.ELLIPSE_RIBBON_2) 
                                    ,new ShapeTypesInfo("LEFT_RIGHT_RIBBON", ShapeTypes.LEFT_RIGHT_RIBBON) 
                                    ,new ShapeTypesInfo("VERTICAL_SCROLL", ShapeTypes.VERTICAL_SCROLL) 
                                    ,new ShapeTypesInfo("HORIZONTAL_SCROLL", ShapeTypes.HORIZONTAL_SCROLL) 
                                    ,new ShapeTypesInfo("WAVE", ShapeTypes.WAVE) 
                                    ,new ShapeTypesInfo("DOUBLE_WAVE", ShapeTypes.DOUBLE_WAVE) 
                                    ,new ShapeTypesInfo("PLUS", ShapeTypes.PLUS) 
                                    ,new ShapeTypesInfo("FLOW_CHART_PROCESS", ShapeTypes.FLOW_CHART_PROCESS) 
                                    ,new ShapeTypesInfo("FLOW_CHART_DECISION", ShapeTypes.FLOW_CHART_DECISION) 
                                    ,new ShapeTypesInfo("FLOW_CHART_INPUT_OUTPUT", ShapeTypes.FLOW_CHART_INPUT_OUTPUT) 
                                    ,new ShapeTypesInfo("FLOW_CHART_PREDEFINED_PROCESS", ShapeTypes.FLOW_CHART_PREDEFINED_PROCESS) 
                                    ,new ShapeTypesInfo("FLOW_CHART_INTERNAL_STORAGE", ShapeTypes.FLOW_CHART_INTERNAL_STORAGE) 
                                    ,new ShapeTypesInfo("FLOW_CHART_DOCUMENT", ShapeTypes.FLOW_CHART_DOCUMENT) 
                                    ,new ShapeTypesInfo("FLOW_CHART_MULTIDOCUMENT", ShapeTypes.FLOW_CHART_MULTIDOCUMENT) 
                                    ,new ShapeTypesInfo("FLOW_CHART_TERMINATOR", ShapeTypes.FLOW_CHART_TERMINATOR) 
                                    ,new ShapeTypesInfo("FLOW_CHART_PREPARATION", ShapeTypes.FLOW_CHART_PREPARATION) 
                                    ,new ShapeTypesInfo("FLOW_CHART_MANUAL_INPUT", ShapeTypes.FLOW_CHART_MANUAL_INPUT) 
                                    ,new ShapeTypesInfo("FLOW_CHART_MANUAL_OPERATION", ShapeTypes.FLOW_CHART_MANUAL_OPERATION) 
                                    ,new ShapeTypesInfo("FLOW_CHART_CONNECTOR", ShapeTypes.FLOW_CHART_CONNECTOR) 
                                    ,new ShapeTypesInfo("FLOW_CHART_PUNCHED_CARD", ShapeTypes.FLOW_CHART_PUNCHED_CARD) 
                                    ,new ShapeTypesInfo("FLOW_CHART_PUNCHED_TAPE", ShapeTypes.FLOW_CHART_PUNCHED_TAPE) 
                                    ,new ShapeTypesInfo("FLOW_CHART_SUMMING_JUNCTION", ShapeTypes.FLOW_CHART_SUMMING_JUNCTION) 
                                    ,new ShapeTypesInfo("FLOW_CHART_OR", ShapeTypes.FLOW_CHART_OR) 
                                    ,new ShapeTypesInfo("FLOW_CHART_COLLATE", ShapeTypes.FLOW_CHART_COLLATE) 
                                    ,new ShapeTypesInfo("FLOW_CHART_SORT", ShapeTypes.FLOW_CHART_SORT) 
                                    ,new ShapeTypesInfo("FLOW_CHART_EXTRACT", ShapeTypes.FLOW_CHART_EXTRACT) 
                                    ,new ShapeTypesInfo("FLOW_CHART_MERGE", ShapeTypes.FLOW_CHART_MERGE) 
                                    ,new ShapeTypesInfo("FLOW_CHART_OFFLINE_STORAGE", ShapeTypes.FLOW_CHART_OFFLINE_STORAGE) 
                                    ,new ShapeTypesInfo("FLOW_CHART_ONLINE_STORAGE", ShapeTypes.FLOW_CHART_ONLINE_STORAGE) 
                                    ,new ShapeTypesInfo("FLOW_CHART_MAGNETIC_TAPE", ShapeTypes.FLOW_CHART_MAGNETIC_TAPE) 
                                    ,new ShapeTypesInfo("FLOW_CHART_MAGNETIC_DISK", ShapeTypes.FLOW_CHART_MAGNETIC_DISK) 
                                    ,new ShapeTypesInfo("FLOW_CHART_MAGNETIC_DRUM", ShapeTypes.FLOW_CHART_MAGNETIC_DRUM) 
                                    ,new ShapeTypesInfo("FLOW_CHART_DISPLAY", ShapeTypes.FLOW_CHART_DISPLAY) 
                                    ,new ShapeTypesInfo("FLOW_CHART_DELAY", ShapeTypes.FLOW_CHART_DELAY) 
                                    ,new ShapeTypesInfo("FLOW_CHART_ALTERNATE_PROCESS", ShapeTypes.FLOW_CHART_ALTERNATE_PROCESS) 
                                    ,new ShapeTypesInfo("FLOW_CHART_OFFPAGE_CONNECTOR", ShapeTypes.FLOW_CHART_OFFPAGE_CONNECTOR) 
                                    ,new ShapeTypesInfo("ACTION_BUTTON_BLANK", ShapeTypes.ACTION_BUTTON_BLANK) 
                                    ,new ShapeTypesInfo("ACTION_BUTTON_HOME", ShapeTypes.ACTION_BUTTON_HOME) 
                                    ,new ShapeTypesInfo("ACTION_BUTTON_HELP", ShapeTypes.ACTION_BUTTON_HELP) 
                                    ,new ShapeTypesInfo("ACTION_BUTTON_INFORMATION", ShapeTypes.ACTION_BUTTON_INFORMATION) 
                                    ,new ShapeTypesInfo("ACTION_BUTTON_FORWARD_NEXT", ShapeTypes.ACTION_BUTTON_FORWARD_NEXT) 
                                    ,new ShapeTypesInfo("ACTION_BUTTON_BACK_PREVIOUS", ShapeTypes.ACTION_BUTTON_BACK_PREVIOUS) 
                                    ,new ShapeTypesInfo("ACTION_BUTTON_END", ShapeTypes.ACTION_BUTTON_END) 
                                    ,new ShapeTypesInfo("ACTION_BUTTON_BEGINNING", ShapeTypes.ACTION_BUTTON_BEGINNING) 
                                    ,new ShapeTypesInfo("ACTION_BUTTON_RETURN", ShapeTypes.ACTION_BUTTON_RETURN) 
                                    ,new ShapeTypesInfo("ACTION_BUTTON_DOCUMENT", ShapeTypes.ACTION_BUTTON_DOCUMENT) 
                                    ,new ShapeTypesInfo("ACTION_BUTTON_SOUND", ShapeTypes.ACTION_BUTTON_SOUND) 
                                    ,new ShapeTypesInfo("ACTION_BUTTON_MOVIE", ShapeTypes.ACTION_BUTTON_MOVIE) 
                                    ,new ShapeTypesInfo("GEAR_6", ShapeTypes.GEAR_6) 
                                    ,new ShapeTypesInfo("GEAR_9", ShapeTypes.GEAR_9) 
                                    ,new ShapeTypesInfo("FUNNEL", ShapeTypes.FUNNEL) 
                                    ,new ShapeTypesInfo("MATH_PLUS", ShapeTypes.MATH_PLUS) 
                                    ,new ShapeTypesInfo("MATH_MINUS", ShapeTypes.MATH_MINUS) 
                                    ,new ShapeTypesInfo("MATH_MULTIPLY", ShapeTypes.MATH_MULTIPLY) 
                                    ,new ShapeTypesInfo("MATH_DIVIDE", ShapeTypes.MATH_DIVIDE) 
                                    ,new ShapeTypesInfo("MATH_EQUAL", ShapeTypes.MATH_EQUAL) 
                                    ,new ShapeTypesInfo("MATH_NOT_EQUAL", ShapeTypes.MATH_NOT_EQUAL) 
                                    ,new ShapeTypesInfo("CORNER_TABS", ShapeTypes.CORNER_TABS) 
                                    ,new ShapeTypesInfo("SQUARE_TABS", ShapeTypes.SQUARE_TABS) 
                                    ,new ShapeTypesInfo("PLAQUE_TABS", ShapeTypes.PLAQUE_TABS) 
                                    ,new ShapeTypesInfo("CHART_X", ShapeTypes.CHART_X) 
                                    ,new ShapeTypesInfo("CHART_STAR", ShapeTypes.CHART_STAR) 
                                    ,new ShapeTypesInfo("CHART_PLUS", ShapeTypes.CHART_PLUS) 
                                    };


  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
 
    // ワークシート生成
    Sheet sheet = workBook.createSheet("Sheet1");
    // 描画元締めの作成
    XSSFDrawing _patr2007 = ((XSSFSheet)sheet).createDrawingPatriarch();
    //int lin = 0;
    short col = 0;
    for (int i=0; i<sti.length; i++) {
      // ClientAnchor生成
      //XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, col, lin, col + 1, lin + 4);
      //XSSFSimpleShape shape = _patr2007.createSimpleShape(anchor);
      if (col != 0) {
        System.out.print(",");
      }
      System.out.print(sti[i]._typeName);
      //shape.setShapeType(sti[i]._typeNo);
      col++;
      if (col > 2) {
        col = 0;
        System.out.println("");
      }
    }
   // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./SetPicture_Book1.xls" : 
                      "./SetPicture_Book1.xlsx");
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
    System.out.println("");
  }
  /** エントリーポイント */

  public static void main(String[] args) {

    if (args.length != 1) {
      System.out.println("エラー：モードを指定して下さい。");
      return;
    }
    else if ( !args[0].equals("2007") ) {
      System.out.println("エラー：モードは2007を指定して下さい。");
      return;
    }
    // 処理の実行
    new SetShapeXSSF().Run(args[0]);

    //System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
