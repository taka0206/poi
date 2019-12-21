// PROGRAM MODULE HELLOWORLD, SAVED Sun Jan 01 16:55:43 JST 2012
100              PROGRAM HELLOWORLD
110              // モード入力
120   L_INPUT:   
130              LINE INPUT "INPUT MODE(X FOR CANCEL) ->",MODE
140              IF MODE="X" OR MODE="x" THEN END
150              IF MODE="2003" OR MODE="2007" THEN GOTO CREATE_WORKBOOK
160              GOTO L_INPUT
170   CREATE_WORKBOOK: 
180              // ワークブックの生成
190              IF MODE="2007" THEN GOTO WORKBOOK_2007
200   WORKBOOK_2003: 
210              LET WORKBOOK=NEW("org.apache.poi.hssf.usermodel.HSSFWorkbook")
220              GOTO MAIN_PROCEDURE
230   WORKBOOK_2007: 
240              LET WORKBOOK=NEW("org.apache.poi.xssf.usermodel.XSSFWorkbook")
250   MAIN_PROCEDURE: 
260              // シートの生成
270              LET SHEET=WORKBOOK->CREATESHEET("HelloWorld")
280              // rowの生成
290              LET ROW=SHEET->CREATEROW(0)
300              // cellの生成
310              LET CELL=ROW->CREATECELL(0)
320              LET ST=WORKBOOK->CREATECELLSTYLE()
330              LET FNT=WORKBOOK->CREATEFONT()
340              CALL FNT->SETFONTNAME("ＭＳ 明朝")
350              LET PNT=NEW("java.lang.Integer",48)
360              CALL FNT->SETFONTHEIGHTINPOINTS(PNT->SHORTVALUE())
370              LET CL=NEW("java.lang.Integer",49)
380              CALL FNT->SETCOLOR(CL->SHORTVALUE())
390              CALL ST->SETFONT(FNT)
400              // cellに値を設定
410              CALL CELL->SETCELLVALUE("Hello POI World On JBasic♪")
420              // ワークブック書き出し
430              IF MODE="2007" THEN GOTO OPEN_OUTFILE_2007
440   OPEN_OUTFILE_2003: 
450              LET FOUT=NEW("java.io.FileOutputStream","./jBasicHello.xls")
460              GOTO WRITE_BOOK
470   OPEN_OUTFILE_2007: 
480              LET FOUT=NEW("java.io.FileOutputStream","./jBasicHello.xlsx")
490   WRITE_BOOK: 
500              CALL WORKBOOK->WRITE(FOUT)
510              CALL FOUT->CLOSE()
520              PRINT  "Done!"
530              END
