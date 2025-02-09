Sub 入力後値貼り付けにする例()

    Dim ws As Worksheet
    Dim 最終行 As Long
    
    ' 対象のシート名を適宜変更してください
    Set ws = ThisWorkbook.Sheets("test")
    
    ' 数式をコピーする最終行（ここでは 5000行目まで）
    最終行 = 10
    
    ' 「原料展開」の数式をセルに表示させるには新しいシート作成し、B1セルに、
    ' =FORMULATEXT(INDIRECT("原料展開!" & ADDRESS(3, ROW(A1), 4)))
    ' と入力し、B1セルをコピーして、B60セルまで貼り付ける。
    ' --- (1) 3行目に数式を入力 ---
    ws.Range("A3").Formula = "=IFERROR(VLOOKUP($AA3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品-部品並び順'!$A:$C,2,FALSE),"""")"
    ws.Range("B3").Formula = "=IFERROR(VLOOKUP($AA3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品-部品並び順'!$A:$C,3,FALSE),"""")"
    ws.Range("C3").Formula = "=IF(B3="","",IF(B3<=10,12+B3*8,15+B3*8))"
    ws.Range("D3").Formula = "=IFERROR(IF(A3="","",VLOOKUP($A3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$EO:$EP,2,FALSE)),0)"
    ws.Range("E3").Formula = "=IFERROR(@INDEX('\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$C:$C,MATCH($D3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$B:$B,0),0),0)"
    ws.Range("F3").Formula = "=IFERROR(VLOOKUP($D3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$B:$EM,$C3+F$1,FALSE),0)"
    ws.Range("G3").Formula = "=IFERROR(IF(LEN(F3)>4,VALUE(LEFT(F3,5)),VALUE(F3)),"****")"
    ws.Range("H3").Formula = "=IF(G3="****","****",IF(LEN(F3)>4,VLOOKUP(F3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]部品マスター'!G:Q,11,FALSE),F3))"
    ws.Range("I3").Formula =  "=IFERROR(VLOOKUP($D3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$B:$EM,$C3+I$1,FALSE),0)"
    ws.Range("J3").Formula =  "=IFERROR(VLOOKUP($D3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$B:$EM,$C3+J$1,FALSE),0)"
    ws.Range("K3").Formula =  "=IFERROR(VLOOKUP($D3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$B:$EM,$C3+K$1,FALSE),0)"
    ws.Range("L3").Formula =  "=IFERROR(VLOOKUP($D3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$B:$EM,$C3+L$1,FALSE),0)"
    ws.Range("M3").Formula =  "=IFERROR(VLOOKUP($D3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$B:$EM,$C3+M$1,FALSE),0)"
    ws.Range("N3").Formula =  "=IFERROR(VLOOKUP($D3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$B:$EM,$C3+N$1,FALSE),0)"
    ws.Range("O3").Formula =  "=IFERROR(VLOOKUP($D3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$B:$EM,$C3+O$1,FALSE),0)"
    ws.Range("P3").Formula =  "=IFERROR(VLOOKUP($D3,出荷数キット!$A:$D,4,FALSE)*M3,0)"
    ws.Range("Q3").Formula =  "=IFERROR(IF(LEN(F3)>4,VLOOKUP(F3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]部品マスター'!$G:$Q,9,FALSE),IFERROR(VLOOKUP(H3,出荷数アイテム!A:M,13,FALSE)+0,N3)),0)"
    ws.Range("R3").Formula =  "=IF(AND(AL3="○",P3>0),IFERROR(IF(LEN(F3)=4,LOOKUP(10^17,LEFT(Q3,COLUMN($1:$1))*1)*P3,(P3*Q3+AH3/M3)/IF(BC3="野菜",VLOOKUP(F3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]部品マスター'!$AM:$BZ,40,FALSE),VLOOKUP(F3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]部品マスター'!$AM:$BH,22,FALSE))),0),IFERROR(IF(LEN(F3)=4,LOOKUP(10^17,LEFT(Q3,COLUMN($1:$1))*1)*P3,P3*Q3/IF(BC3="野菜",VLOOKUP(F3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]部品マスター'!$AM:$BZ,40,FALSE),VLOOKUP(F3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]部品マスター'!$AM:$BH,22,FALSE))),0))"
    ws.Range("S3").Formula =  "=IFERROR(@INDEX('\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$FA:$FA,MATCH($D3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$EP:$EP,0),0),"")"
    ws.Range("T3").Formula =  "=IFERROR(@INDEX('\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[取引先マスター.xlsm]取引先マスター'!$C:$C,MATCH($D3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[取引先マスター.xlsm]取引先マスター'!$A:$A,0),0),0)"
    ws.Range("U3").Formula =  "=IF(V3="","",IF(LEN(V3)=4,COUNTIF(U$2:U2,"<1000")+1,IF(LEN(V3)=5,COUNTIF(V$3:V3,"*×*")+1000,"")))"
    ws.Range("V3").Formula =  "=IF(S3="○",IF(AND(LEN(F3)=4,COUNTIFS(F$2:F3,F3,S$2:S3,"<>"&"×")=1),F3,""),IF(AND(LEN(F3)=4,COUNTIFS(F$2:F3,F3,S$2:S3,"<>"&"○")=1,COUNTIFS(F:F,F3,S:S,"○")=0),"×"&F3,""))"
    ws.Range("W3").Formula =  "=IFERROR(IF(COUNTIF($D$3:$D3,$D3)=1,VLOOKUP($D3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$B:$T,19,FALSE),0),0)"
    ws.Range("X3").Formula =  "=IFERROR(IF(COUNTIF($AH$1:$AI$1,$H3)>0,$H3,IF(COUNTIF(ポリ袋!$D:$D,$H3)>0,$J3,VLOOKUP($F3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]部品マスター'!G:$AA,21,FALSE))),0)"
    ws.Range("Y3").Formula =  "=IFERROR(P3*Q3,0)"
    ws.Range("Z3").Formula =  "=IFERROR(VLOOKUP($D3,受注入力!$A:$D,4,FALSE)*VLOOKUP($D3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$EP:$FP,Z$1,FALSE),0)"
    ws.Range("AA3").Formula = "=ROW()-2"
    ws.Range("AB3").Formula = "=IF(T3="コープデリ","コープデリ",IF(T3="ユーコープ","ユーコープ",IF(T3="ピースデリ","ピースデリ","業務用")))"
    ws.Range("AC3").Formula = "=I3&T3"
    ws.Range("AD3").Formula = "=IF(COUNTIF(小分け品コード!$C:$C,原料展開!$F3)>0,"","×")"
    ws.Range("AE3").Formula = "=INDEX('\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]部品マスター'!$N:$N,MATCH($F3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]部品マスター'!$G:$G,0))"
    ws.Range("AF3").Formula = "=IF($AE3=$O3,"○","X")"
    ws.Range("AG3").Formula = "=IFERROR(INDEX(出荷数キット!$E:$E,MATCH(原料展開!$D3,出荷数キット!$A:$A,0)),0)"
    ws.Range("AH3").Formula = "=LOOKUP(10^17,LEFT(Q3,COLUMN($1:$1))*1)*M3"
    ws.Range("AI3").Formula = "=INDEX('\\Afnewt320-kyoyu\社内共有\AFSKS\生産管理\[原料マスター.xlsm]原料マスター'!$X:$X,MATCH(H3,'\\Afnewt320-kyoyu\社内共有\AFSKS\生産管理\[原料マスター.xlsm]原料マスター'!$A:$A,0))"
    ws.Range("AJ3").Formula = "=IF(AND(H3>8000,H3<9999,H3<>8453),1,LOOKUP(10^17,LEFT(AI3,COLUMN($1:$1))*1))"
    ws.Range("AK3").Formula = "=IFERROR((AH3/AJ3)*AG3,0)"
    ws.Range("AL3").Formula = "=IF(COUNTIF(AM3:AN3,"○")>=1,"○","X")"
    ws.Range("AM3").Formula = "=IFERROR(IF(AND(T3="イトーヨーカドー",INDEX(W:W,MATCH(A3,A:A,0))="ばんじゅう用内袋№110 ナチュラル",O3="生産"),"○","X"),"X")"
    ws.Range("AN3").Formula = "=IF(OR($W3=0,COUNTIF($AJ$1:$AL$1,$W3)>0),0,COUNTIFS($W$3:$W3,"<>"&$AJ$1,$W$3:$W3,"<>"&$AK$1,$W$3:$W3,"<>"&$AL$1,$W$3:$W3,"<>0"))"
    ws.Range("AO3").Formula = "=IF(AND(COUNTIF($AH$1:$AI$1,$H3)=0,OR($X3=0,COUNTIF($AJ$1:$AL$1,$X3)>0,$AD3="")),0,COUNTIFS($X$3:$X3,"<>"&$AJ$1,$X$3:$X3,"<>"&$AK$1,$X$3:$X3,"<>"&$AL$1,$X$3:$X3,"<>0",$AD$3:$AD3,"×")+$AM$1)"
    ws.Range("AP3").Formula = "=SUMIFS($Y:$Y,$G:$G,$G3,$O:$O,"生産",$AB:$AB,"業務用",$AD:$AD,"×")"
    ws.Range("AQ3").Formula = "=SUMIFS($Y:$Y,$G:$G,$G3,$O:$O,"生産",$AB:$AB,"ユーコープ",$AD:$AD,"×")"
    ws.Range("AR3").Formula = "=SUMIFS($Y:$Y,$G:$G,$G3,$O:$O,"生産",$AB:$AB,"コープデリ",$AD:$AD,"×")"
    ws.Range("AS3").Formula = "=SUMIFS($Y:$Y,$G:$G,$G3,$O:$O,"生産",$AB:$AB,"ピースデリ",$AD:$AD,"×")"
    ws.Range("AT3").Formula = "=IFERROR(IF(IF(LEN(F3)>=5,INDEX('\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]部品マスター'!$Y:$Y,MATCH($F3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]部品マスター'!$G:$G,0)),0)>0,G3,0),0)"
    ws.Range("AU3").Formula = "=IF(OR(AT3="",AT3=0),"",COUNTIF($AT$3:AT3,AT3))"
    ws.Range("AV3").Formula = "=IF(AU3=1,COUNTIF($AU$3:AU3,1),"")"
    ws.Range("AW3").Formula = "=IFERROR(INDEX('\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$FA:$FA,MATCH(D3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]商品マスター'!$B:$B,0)),"")"
    ws.Range("AX3").Formula = "=IFERROR(INDEX('\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]部品マスター'!$I:$I,MATCH($F3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]部品マスター'!$G:$G,0)),"")"
    ws.Range("AY3").Formula = "=IFS(AY2="取引先No",1,T2=T3,AY2,TRUE,AY2+1)"
    ws.Range("AZ3").Formula = "=IFERROR(INDEX('\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]部品マスター'!$BH:$BH,MATCH($F3,'\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\[販売商品マスター.xlsm]部品マスター'!$G:$G,0)),0)"
    ws.Range("BA3").Formula = "=IFERROR((Y3/AZ3)/1000,0)"
    ws.Range("BB3").Formula = "=IFERROR(IF(OR(LEN(F3)=9,LEN(F3)=7),LEFT(F3,3),""),"")"
    ws.Range("BC3").Formula = "=IFERROR(INDEX('\\Afnewt320-kyoyu\社内共有\AFSKS\生産管理\[原料マスター.xlsm]原料マスター'!$AC:$AC,MATCH($H3,'\\Afnewt320-kyoyu\社内共有\AFSKS\生産管理\[原料マスター.xlsm]原料マスター'!$A:$A,0)),"")"
    ws.Range("BD3").Formula = "=IFERROR(INDEX(出荷順!$A:$A,MATCH(原料展開!$T3,出荷順!$C:$C,0)),"-")"



    
    
    ' --- (2) 3行目から最終行までオートフィルでコピー ---
    '   ここでは列A～BDまで一括コピーする例
    ws.Range("A3:BD3").AutoFill Destination:=ws.Range("A3:BD" & 最終行)
    
    ' --- (3) 計算式を値貼り付けに変更 ---
    With ws.Range("A3:BD" & 最終行)
        .Value = .Value
    End With
    
    MsgBox "完了しました。"

End Sub

