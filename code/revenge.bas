// PROGRAM MODULE HELLOWORLD2, SAVED Sun Jan 01 16:57:03 JST 2012
100              PROGRAM HELLOWORLD2
110              // ƒ‰ƒbƒp[‚Ì¶¬
120              LET WRAPPER=NEW("JBasWrapper")
130   L_INPUT:   
140              LINE INPUT "INPUT MODE(X FOR CANCEL) ->",MODE
150              IF MODE="X" OR MODE="x" THEN END
160              IF MODE="2003" OR MODE="2007" THEN GOTO CREATE_WORKBOOK
170              GOTO L_INPUT
180   CREATE_WORKBOOK: 
190              CALL WRAPPER->CREATEWORKSHEETANDROWANDCELL(MODE,"HelloWorld")
200              CALL WRAPPER->SETFONTANDSTYLE("‚l‚r –¾’©",48,49)
210              CALL WRAPPER->SETCELLVALUE("Hello World On JBasicô")
220              CALL WRAPPER->WRITE("./HelloWorld-JBasic")
230              PRINT  "Done!"
240              END
