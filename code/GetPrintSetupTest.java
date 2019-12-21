import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

/**
 * プリンター設定取得テスト
 */
public class GetPrintSetupTest {
  /**
   * 用紙サイズ文字列取得
   *@pSize 用紙サイズ番号
   */
  protected String deCodePaparSize(short pSize) {
    switch (pSize) {
      case PrintSetup.A3_PAPERSIZE :
        return "A3 - 297x420 mm";
      case PrintSetup.A4_EXTRA_PAPERSIZE :
        return "A4 Extra - 9.27 x 12.69 in";
      case PrintSetup.A4_PAPERSIZE :
        return "A4 - 210x297 mm";
      case PrintSetup.A4_PLUS_PAPERSIZE :
        return "A4 Plus - 210x330 mm";
      case PrintSetup.A4_ROTATED_PAPERSIZE :
        return "A4 Rotated - 297x210 mm";
      case PrintSetup.A4_SMALL_PAPERSIZE :
        return "A4 Small - 210x297 mm";
      case PrintSetup.A4_TRANSVERSE_PAPERSIZE :
        return "A4 Transverse - 210x297 mm";
      case PrintSetup.A5_PAPERSIZE :
        return "A5 - 148x210 mm";
      case PrintSetup.B4_PAPERSIZE :
        return "B4 (JIS) 250x354 mm";
      case PrintSetup.B5_PAPERSIZE :
        return "B5 (JIS) 182x257 mm";
      case PrintSetup.ELEVEN_BY_SEVENTEEN_PAPERSIZE :
        return "11 x 17 in";
      case PrintSetup.ENVELOPE_10_PAPERSIZE :
        return "US Envelope #10 4 1/8 x 9 1/2";
      case PrintSetup.ENVELOPE_9_PAPERSIZE :
        return "US Envelope #9 3 7/8 x 8 7/8";
      case PrintSetup.ENVELOPE_C3_PAPERSIZE :
        return "Envelope C3 324x458 mm";
      case PrintSetup.ENVELOPE_C4_PAPERSIZE :
        return "Envelope C4 229x324 mm";
      case PrintSetup.ENVELOPE_C5_PAPERSIZE :
        return "Envelope C5";
      case PrintSetup.ENVELOPE_C6_PAPERSIZE :
        return "Envelope C6 114x162 mm";
      case PrintSetup.ENVELOPE_DL_PAPERSIZE :
        return "Envelope DL 110x220 mm";
      case PrintSetup.ENVELOPE_MONARCH_PAPERSIZE :
        return "Envelope Nonarch";
      case PrintSetup.EXECUTIVE_PAPERSIZE :
        return "US Executive 7 1/4 x 10 1/2 in";
      case PrintSetup.FOLIO8_PAPERSIZE :
        return "Folio 8 1/2 x 13 in";
      case PrintSetup.LEDGER_PAPERSIZE :
        return "US Ledger 17 x 11 in";
      case PrintSetup.LEGAL_PAPERSIZE :
        return "US Legal 8 1/2 x 14 in";
      case PrintSetup.LETTER_PAPERSIZE :
        return "US Letter 8 1/2 x 11 in";
      case PrintSetup.LETTER_ROTATED_PAPERSIZE :
        return "US Letter Rotated 11 x 8 1/2 in";
      case PrintSetup.LETTER_SMALL_PAGESIZE :
        return "US Letter Small 8 1/2 x 11 in";
      case PrintSetup.NOTE8_PAPERSIZE :
        return "US Note 8 1/2 x 11 in";
      case PrintSetup.QUARTO_PAPERSIZE :
        return "Quarto 215x275 mm";
      case PrintSetup.STATEMENT_PAPERSIZE :
        return "US Statement 5 1/2 x 8 1/2 in";
      case PrintSetup.TABLOID_PAPERSIZE :
        return "US Tabloid 11 x 17 in";
      case PrintSetup.TEN_BY_FOURTEEN_PAPERSIZE :
        return "10 x 14 in";
    }
    return "unknown";
  }
  /**
   * プリンター設定出力
   *@psetup PrintSetupの参照
   */
  protected void PrintPrintSettings(PrintSetup psetup) {
    System.out.println("印刷部数                 : " + 
      psetup.getCopies());
    System.out.println("FitHeight                : " + 
      psetup.getFitHeight());
    System.out.println("FitWidth                 : " + 
      psetup.getFitWidth());
    System.out.println("フッター余白             : " + 
      psetup.getFooterMargin());
    System.out.println("ヘッダー余白             : " + 
      psetup.getHeaderMargin());
    System.out.println("水平解像度               : " + 
      psetup.getHResolution());
    System.out.println("LeftToRight              : " + 
      psetup.getLeftToRight());
    System.out.println("ランドスケープモード     : " + 
      psetup.getLandscape());
    System.out.println("白黒モード               : " + 
      psetup.getNoColor());
    System.out.println("NoOrientation            : " + 
      psetup.getNoOrientation());
    System.out.println("セル内コメント印刷モード : " + 
      psetup.getNotes());
    System.out.println("PageStart                : " + 
      psetup.getPageStart());
    System.out.println("用紙サイズ               : " + 
      deCodePaparSize(psetup.getPaperSize()));
    System.out.println("印刷倍率                 : " + 
      psetup.getScale());
    System.out.println("ページ番号印刷           : " + 
      psetup.getUsePage());
    System.out.println("ValidSettings            : " + 
      psetup.getValidSettings());
    System.out.println("垂直解像度               : " + 
      psetup.getVResolution());
  }
  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // シートの生成
    Sheet sheet = workBook.createSheet();
    // プリンター設定の取得
    PrintSetup psetup = sheet.getPrintSetup(); 
    // 各設定を出力する。
    PrintPrintSettings(psetup);
  }
  /** エントリーポイント */
  public static void main(String[] args) {
    if (args.length != 1) {
      System.out.println("エラー：モードを指定して下さい。");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    // 処理の実行
    new GetPrintSetupTest().Run(args[0]);
    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
