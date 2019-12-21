import java.io.*;
import org.apache.poi.ss.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.*;
/**
 * サンプルランチャーシート作成
 */
class buildLauncherSheet {

	protected String _mode;		// 動作モード
	protected Workbook _workBook;	// ランチャーワークブックのインスタンス
	protected int _classPos;		// クラス名の桁位置
	protected String[] _breakKeys; // キーブレーク見出退避領域
	protected int[] _breakPos; // キーブレーク行番号退避領域
	/** 
	 * 初期処理
	 */
	protected boolean init() {
    FileInputStream fis = null;
    // ワークブックを読み込む
    _workBook = null;
    try {
      fis = new FileInputStream( _mode.equals("2003") ? "./SampleLauncherORG.xls" : "./SampleLauncherORG.xlsx");
      _workBook = _mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" + e.toString());
      return false;
    }
		// クラス名の桁位置を取得
		_classPos = (int)_workBook.getSheet("データシート").getRow(0).getCell(1).getNumericCellValue();
		System.out.println("codePos = " + _classPos);
		// キーブレーク見出退避領域を準備する。
		_breakKeys = new String[_classPos - 1];
		// キーブレーク見出退避領域初期化
		for (int i=0; i<_classPos - 1; i++) {
			_breakKeys[i] = "";
		}
		// キーブレイク行番号退避領域を準備する。
		_breakPos = new int[_classPos - 1];
		// キーブレーク行番号退避領域初期化
		for (int i=0; i<_classPos - 1; i++) {
			_breakPos[i] = -1;
		}
		
		return true;
	}
	/** 
	 * ランチャーシート作成処理
	 */
	protected void buildSheet() {
		CreationHelper cHelper = _workBook.getCreationHelper();
		// リンクセル用スタイルを作成しておく
		CellStyle style = _workBook.createCellStyle();
		Font fnt = _workBook.createFont();
		fnt.setFontName("ＭＳ ゴシック");
		fnt.setFontHeightInPoints((short)9);
		fnt.setColor((short)org.apache.poi.hssf.util.HSSFColor.BLUE.index);
		fnt.setUnderline(Font.U_SINGLE);
		style.setFont(fnt);
		// データシートとランチャーシートの取得
		Sheet dSheet = _workBook.getSheet("データシート");
		Sheet lSheet = _workBook.getSheet("ランチャーシート");
		// データシート、ランチャーシートとも3行目から処理
    for (int i=2; i<=dSheet.getLastRowNum(); i++) {
				// データシートからRowの取得
				Row dRow = dSheet.getRow(i);
				// ランチャーシートに行生成
				Row lRow = lSheet.createRow(i);
			// 列を処理
			for (int j=0; j<_classPos; j++) {
				Cell cell = dRow.getCell(j);
				if (cell != null) {
					String s = cell.getStringCellValue();
					System.out.println(s);
					if (s.equals(_breakKeys[j]) == false) {
						// 見出がブレークすればランチャーシートに設定
						lRow.createCell(j).setCellValue(s);
						System.out.println("ランチャーシートに項目設定");
						if (_breakPos[j] != -1) {
							if ( (i - _breakPos[j]) > 1 ) {
								// 間が開いている場合Cellを縦にマージする。
								lSheet.addMergedRegion(new CellRangeAddress(_breakPos[j],i-1,j,j));
							}
							
						}
						_breakPos[j] = i;	// キーブレーク行番号に現在の行を設定
					}
					_breakKeys[j] = s;
				}
			}
			// クラス名関連処理
			Cell cell = dRow.getCell(_classPos);
			if (cell != null) {
				String className = cell.getStringCellValue();
				if (className.equals("") == false) {
					lRow.createCell(_classPos).setCellValue(className);
					boolean bBook1 = dRow.getCell(_classPos + 1).getBooleanCellValue();	// Book生成フラグ
					lRow.createCell(_classPos + 1).setCellValue(bBook1);
					boolean both = dRow.getCell(_classPos + 2).getBooleanCellValue();
					Cell fCell = lRow.createCell(_classPos + 2);
					fCell.setCellValue("ソースファイル参照");
					fCell.setCellStyle(style);
					Hyperlink fLink = cHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
					fLink.setAddress("");
					fCell.setHyperlink(fLink);
					Cell bCell = lRow.createCell(_classPos + 3);
					bCell.setCellValue("ビルド");
					bCell.setCellStyle(style);
					Hyperlink bLink = cHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
					bLink.setAddress("");
					bCell.setHyperlink(bLink);
					Cell exCell2003 = lRow.createCell(_classPos + 4);
					exCell2003.setCellValue("実行(2003)");
					exCell2003.setCellStyle(style);
					Hyperlink ex3Link = cHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
					ex3Link.setAddress("");
					exCell2003.setHyperlink(ex3Link);
					if (both) {
						Cell exCell2007 = lRow.createCell(_classPos + 5);
						exCell2007.setCellValue("実行(2007)");
						exCell2007.setCellStyle(style);
						Hyperlink ex7Link = cHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
						ex7Link.setAddress("");
						exCell2007.setHyperlink(ex7Link);
					}
				}
			}
		}
		// 最後のセルマージを行う。
		for (int i=0;i<_classPos-1; i++) {
			if (_breakPos[i] != -1 && _breakPos[i] != lSheet.getLastRowNum()) {
				lSheet.addMergedRegion(new CellRangeAddress(_breakPos[i],lSheet.getLastRowNum(),i,i));
			}
		}
		// Book出力カラムを非表示に
		lSheet.setColumnHidden(_classPos + 1, true);
		// 以後のカラムを自動幅設定にする。
		for (int i=_classPos + 1; i<_classPos + 6; i++) {
			lSheet.autoSizeColumn(i);
		}
		// シートを分割する。
		lSheet.createFreezePane(_classPos + 2, 2);
		// ランチャーシートの構築が終わればデータシートを削除する。
		_workBook.removeSheetAt(_workBook.getSheetIndex("データシート"));
		// 作業用シートを非表示にする。
		_workBook.setSheetHidden(_workBook.getSheetIndex("作業用シート"), true);
	}
	/**
	 * Excelブック出力処理
	 */
	protected void write() {
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( _mode.equals("2003") ? "./Book1.xls" : 
                      "./Book1.xlsx");
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
  /** 処理の実行
   * @param モード
   */
  public void Run(String mode) {

		_mode = mode;
		// 初期処理
		if (init() == false) {
			return;
		}
		// ランチャーシート作成
		buildSheet();
    // ワークブック書き出し
		write();
	

    System.out.println("done!");
  }
  /** エントリーポイント */
  public static void main(String[] args) {
    if (args.length != 1) {
      System.out.println("エラー：モードを指定してください。");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    new buildLauncherSheet().Run(args[0]);
		System.out.print("リターンキーで終了……");
		try {
			int c = System.in.read();
		}
		catch (Exception e) {
		}
  }
}
