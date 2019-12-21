import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * ハイパーリンク解除のテスト
 */
public class RemoveLinkTest {

  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    FileInputStream fis = null;

    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // シートの生成 Sheet0
    Sheet sheet = workBook.createSheet();
    // Sheet1
    Sheet sheet2 = workBook.createSheet();
    // RowとCellの生成
    for (int i=0; i<5; i++){
      sheet.createRow(i);
    }
    // Sheet1にRowとCellを生成し値設定
    sheet2.createRow(0).createCell(0).setCellValue("リンク先");
    
    CreationHelper cHelper = workBook.getCreationHelper();
    // ハイパーリンク設定
    // URL
    Cell cellUrl = sheet.getRow(0).createCell(0);
    Hyperlink linkURL = cHelper.createHyperlink(Hyperlink.LINK_URL);
    linkURL.setAddress("http://poi.apache.org/");
    //linkURL.setAddress("");
    cellUrl.setCellValue("ポポイッとPOI");
    cellUrl.setHyperlink(linkURL);
    // Document
    Cell cellDoc = sheet.getRow(1).createCell(0);
    Hyperlink linkDoc = cHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
    linkDoc.setAddress("SHeet1!A1");
    //linkDoc.setAddress("");
    cellDoc.setCellValue("Sheet1へ");
    cellDoc.setHyperlink(linkDoc);
    // Mail
    Cell cellEMail = sheet.getRow(2).createCell(0);
    Hyperlink linkMail = cHelper.createHyperlink(Hyperlink.LINK_EMAIL);
    linkMail.setAddress("mailto:impl_person@yahoo.co.jp");
    //linkMail.setAddress("");
    cellEMail.setCellValue("おたよりはこちら");
    cellEMail.setHyperlink(linkMail);
    // File
    Cell cellFile = sheet.getRow(3).createCell(0);
    Hyperlink linkFile = cHelper.createHyperlink(Hyperlink.LINK_FILE);
    linkFile.setAddress("Book1.xlsx");
    //linkFile.setAddress("");
    cellFile.setCellValue("別のブック");
    cellFile.setHyperlink(linkFile);
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./hlink.xls" : 
                      "./hlink.xlsx");
      workBook.write(out);
    }catch(IOException e){
      System.out.println("ブックの書き込みに失敗しました。\n" + e.toString());
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println("ブックの書き込みに失敗しました。\n" + e.toString());
      }
    }
    System.out.println("done!");
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
    new RemoveLinkTest().Run(args[0]);
  }
}
