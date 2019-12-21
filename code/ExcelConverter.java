import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * Excelブックコンバーター
 * 2003以前<->2007以降 相互変換
 */
class ExcelConverter {
	/**
	 * エントリーポイント
	 *@param args[0] 動作モード u = 2003->2007 d = 2007->2003
	 *@param 入力ワークブックファイル名
	 *@param 出力ワークブックファイル名
	 */
	public static void main(String args[]) {
		// パラメーターチェック
		if (args.length != 3) {
			System.out.println("パラメーターエラーです。");
			return;
		}
		String mode = args[0];
		if (!mode.equals("u") && !mode.equals("d")) {
			System.out.println("動作モードは d または u で指定します。");
			return;
		}
		// 処理開始
		// Excelブックをオープン
		Workbook workBook = null;
		try {
			if (mode.equals("u") {
				workBook = new HSSFWorkbook(new FileInputStream(args[1]));
			}
			else {
				workBook = new XSSFWorkbook(new FileInputStream(args[1]));
			}
		}
		catch (Exception e) {
			System.out.println("入力ブックが開けません。" + e.ToStrig());
		}
		
	}
}

