'========================================================
' X-Projects v0.4.0 - Copyright (C) 2011 M.Nomura
'========================================================
' 変更履歴
'  v0.1.0 初期バージョン
'  v0.2.0 日付計算しない担当を追加
'  v0.2.1 回答列を追加、名前を整理
'  v0.2.2 設定シートに進捗情報欄を追加
'  v0.3.0 CCPM対応、担当者状況と予定工数履歴を追加
'  v0.3.1 CCPMじゃない進捗管理もバックポート
'  v0.3.2 予定工数履歴に実績・残り・予定工数計を追加
'  v0.3.3 進捗推移グラフの出力を追加
'  v0.3.4 チケット削除・追加を考慮した予定工数履歴の更新
'  v0.3.5 進捗報告日確認を追加
'  v0.4.0 バッファ管理機能の追加
'========================================================

'設定シート名
Const cCfgShtName = "設定"
Const cYkrShtName = "予定工数履歴"

Dim nDayKosu       '１日と計算する工数
Dim nDefKosu       '未入力時の予定工数
Dim sHolidayOfWeek '休日の曜日
Dim sReCalcDate    '再計算する開始日付

'CSV作成ボタン押下時
Private Sub btnMakeCsv_Click()

    Dim file_source As Object
    Dim file_target As Object

    csv_file_name = ActiveWorkbook.Path + "\" + ActiveSheet.Name + ".csv"

    code_source = "Shift_JIS"
    code_target = "UTF-8"

    char_source = ","
    char_target = ";"

    Application.DisplayAlerts = False
   
' 仮シートを複製
    ActiveSheet.Copy

' CSV形式で保存
    ActiveWorkbook.SaveAs Filename:=csv_file_name, FileFormat:=xlCSV, _
        CreateBackup:=False

' 仮シートを閉じる
    ActiveWindow.Close Savechanges:=False
   

' **********************************
'      CSVファイルをUTF-8に変換
' **********************************
    
' ADODB.Streamを参照
    Set file_source = CreateObject("ADODB.Stream")

' CSVファイルの読み込み
    With file_source
        .Charset = code_source
        .Open
        .LoadFromFile csv_file_name
        char_temp = .ReadText
    End With

' 置換処理
'    char_temp = Replace(char_temp, char_source, char_target)

' CSVファイルの書き出し
    Set file_target = CreateObject("ADODB.Stream")
    With file_target
        .Charset = code_target
        .Open
        .WriteText char_temp
    End With
    
' 文字コードの変換
    file_source.copyto file_target
    file_target.savetofile csv_file_name, 2

End Sub

'日付再計算ボタン押下時
Private Sub btnCalcDate_Click()

    '担当者ごと、順、№でソートして
    '最初の開始日から休日を考慮して
    '予定工数から割り出した日数で
    '開始日、期日を自動計算する
    
    '進捗報告日確認・セット
    sRDate = Format(Now, "yyyy/mm/dd")
    If Sheets(cCfgShtName).Range("進捗報告日確認").Value = "確認する" Then
        sRet = InputBox("指定した進捗報告日で日付再計算処理を実行しますか？", "進捗報告日確認", sRDate)
        If sRet = "" Then
            Exit Sub
        End If
        sRDate = sRet
    End If
    Sheets(cCfgShtName).Range("進捗報告日").Value = sRDate
    Sheets(cCfgShtName).Range("進捗報告日2").Value = sRDate
    
    '設定 読み込み
    nDayKosu = Sheets(cCfgShtName).Range("工数１日")
    nDefKosu = Sheets(cCfgShtName).Range("工数未入力")
    sHolidayOfWeek = ""
    If Sheets(cCfgShtName).Range("休日曜日").Rows(1) <> "" Then
        sHolidayOfWeek = sHolidayOfWeek & "2"
    End If
    If Sheets(cCfgShtName).Range("休日曜日").Rows(2) <> "" Then
        sHolidayOfWeek = sHolidayOfWeek & "3"
    End If
    If Sheets(cCfgShtName).Range("休日曜日").Rows(3) <> "" Then
        sHolidayOfWeek = sHolidayOfWeek & "4"
    End If
    If Sheets(cCfgShtName).Range("休日曜日").Rows(4) <> "" Then
        sHolidayOfWeek = sHolidayOfWeek & "5"
    End If
    If Sheets(cCfgShtName).Range("休日曜日").Rows(5) <> "" Then
        sHolidayOfWeek = sHolidayOfWeek & "6"
    End If
    If Sheets(cCfgShtName).Range("休日曜日").Rows(6) <> "" Then
        sHolidayOfWeek = sHolidayOfWeek & "7"
    End If
    If Sheets(cCfgShtName).Range("休日曜日").Rows(7) <> "" Then
        sHolidayOfWeek = sHolidayOfWeek & "1"
    End If
    sReCalcDate = Sheets(cCfgShtName).Range("再計算開始日付")
    
    'MsgBox WorkDateAdd("2011/10/10", 0)
    'Exit Sub
    
    'Application.ScreenUpdating = False
    
    '担当者、順、№でソート
    Range("データ").Sort _
        Key1:=Range("担当者"), Order1:=xlAscending, _
        Key2:=Range("順"), Order2:=xlAscending, _
        Key3:=Range("No"), Order3:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin, _
        DataOption1:=xlSortNormal, DataOption2:=xlSortTextAsNumbers, DataOption3:=xlSortNormal
    
    '最終行、各カラム位置取得
    nLastRow = Cells.SpecialCells(xlCellTypeLastCell).Row
    nTantoCol = Range("担当者").Column
    nSDateCol = Range("開始日").Column
    nEDateCol = Range("期日").Column
    nTKosuCol = Range("当初工数").Column
    nYKosuCol = Range("予定工数").Column
    
    nTKosuColGT = 0
    nYKosuColGT = 0
    nJKosuColGT = 0
    sDateGT = ""
    
    '担当者リスト取得
    Set oTantos = CreateObject("Scripting.Dictionary")
    
    arTantos = Range(Cells(2, nTantoCol), Cells(nLastRow, nTantoCol))
    
    For Each sTanto In arTantos
        If Not oTantos.Exists(sTanto) And Trim(sTanto) <> "" Then
            oTantos.Add sTanto, Null
        End If
    Next
    
    'デフォルトの開始日付：現在日＋１
    sDefDate = Format(Now + 1, "yyyy/mm/dd")
    
    '担当者状況のクリア
    nTjOff = 0
    With Sheets(cCfgShtName)
        With .Range(.Range("担当者状況雛型行").Offset(1, 0), .Range("担当者状況雛型行").Offset(50, 0))
            .Clear
        End With
    End With
    
    '担当者ごとにループ
    Dim oFind
    For Each sTanto In oTantos
        'Debug.Print sTanto
        sDate = ""
        nKosu = 0
        
        nTKosuColST = 0
        nYKosuColST = 0
        nJKosuColST = 0
        
        '対象外判定
        Set oFind = Sheets(cCfgShtName).Range("日付計算対象外").Find(sTanto, , xlFormulas, xlWhole)
        
        'チケット行ごとにループ
        For i = 2 To nLastRow
            
            '処理中の担当者かどうか
            If Cells(i, nTantoCol) = sTanto Then
                
                '日付再計算する担当者かどうか
                If oFind Is Nothing Then
                    
                    '担当の開始かどうか
                    If sDate = "" Then
                        '再計算かどうか
                        If sReCalcDate = "" Then
                            '開始日が未設定
                            If Cells(i, nSDateCol) = "" Then
                                sDefDate = InputBox("開始日が未設定：" & sTanto, "日付計算", sDefDate)
                                If sDefDate = "" Then
                                    GoTo LOOP_EXIT
                                End If
                                sDate = sDefDate
                            Else
                                sDate = Cells(i, nSDateCol)
                            End If
                            
                            '開始日を営業日にする
                            sDate = WorkDateAdd(sDate, 0)
                            
                             Cells(i, nSDateCol) = sDate
                        Else
                            '開始日が未設定は無視
                            If Cells(i, nSDateCol) <> "" Then
                                '再計算開始日以上なら
                                If CDate(sReCalcDate) <= CDate(Cells(i, nSDateCol)) Then
                                    sDate = Cells(i, nSDateCol)
                                End If
                            End If
                        End If
                    Else
                        Cells(i, nSDateCol) = sDate
                    End If
                    
                    '計算開始中かどうか
                    If sDate <> "" Then
                        
                        '予定工数が未設定
                        If Cells(i, nYKosuCol) = "" Then
                            Cells(i, nYKosuCol) = nDefKosu
                        End If
                        
                        '当初工数が未設定
                        If Cells(i, nTKosuCol) = "" Then
                            Cells(i, nTKosuCol) = Cells(i, nYKosuCol)
                        End If
                        
                        '日数と余り工数
                        nTKosu = Cells(i, nTKosuCol)
                        nYKosu = Cells(i, nYKosuCol)
                        nAKosu = nKosu + nYKosu
                        nDay = Int((nAKosu - 1) / nDayKosu)
                        nKosu = nAKosu Mod nDayKosu
                        
                        '開始日
                        sSDate = sDate
                        
                        '期日の設定
                        sDate = WorkDateAdd(sDate, nDay)
                        Cells(i, nEDateCol) = sDate
                        sEDate = sDate
                        
                        '次の開始日を計算
                        If nYKosu <> 0 And nKosu = 0 Then
                            sDate = WorkDateAdd(sDate, 1)
                        End If
                    End If
                
                Else '日付再計算しない担当者の場合
                    
                    '予定工数
                    nYKosu = Cells(i, nYKosuCol)
                    
                    '当初工数
                    nTKosu = Cells(i, nTKosuCol)
                    
                    '開始日
                    sSDate = Cells(i, nSDateCol)
                    
                    '期日
                    sEDate = Cells(i, nEDateCol)
                    
                End If
                        
                '担当者計、総合計の計算
                nTKosuColST = nTKosuColST + nTKosu
                nYKosuColST = nYKosuColST + nYKosu
                nTKosuColGT = nTKosuColGT + nTKosu
                nYKosuColGT = nYKosuColGT + nYKosu
                
                '実績工数計
                If sEDate < sRDate Then
                    nJKosuColST = nJKosuColST + nYKosu
                    nJKosuColGT = nJKosuColGT + nYKosu
                ElseIf sSDate <= sRDate Then
                    nTmpKosu = (DateDiff("d", CDate(sSDate), CDate(sRDate)) + 1) * nDayKosu
                    nJKosuColST = nJKosuColST + nTmpKosu
                    nJKosuColGT = nJKosuColGT + nTmpKosu
                End If
                
            End If
            
        Next
      
        '担当者状況の出力
        'Debug.Print sTanto & ", " & sEDate & ", " & nTKosuColST & ", " & nYKosuColST
        With Sheets(cCfgShtName).Range("担当者状況雛型行")
            '雛型行コピー
            If nTjOff > 0 Then
                .Copy .Offset(nTjOff, 0)
            End If
            '行情報セット
            With .Offset(nTjOff, 0)
                .Value = Array(sTanto, sEDate, nTKosuColST, nJKosuColST, nYKosuColST - nJKosuColST, nYKosuColST, "", "", "")
                .Cells(1, 7).FormulaR1C1 = "=RC[-1]/工数１日"
                .Cells(1, 8).FormulaR1C1 = "=RC[-1]/20"
                .Cells(1, 9).FormulaR1C1 = "=(RC[-3]/RC[-6]-1)*100"
            End With
            nTjOff = nTjOff + 1
        End With
        
        If sDateGT < sEDate Then sDateGT = sEDate
        
    Next

LOOP_EXIT:
    
    Range("データ").Sort _
        Key1:=Range("No"), Order1:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin, _
        DataOption1:=xlSortNormal, DataOption2:=xlSortTextAsNumbers, DataOption3:=xlSortNormal
 
    'プロジェクト状況の出力
    'Debug.Print "合計" & ", " & sDateGT & ", " & nTKosuColGT & ", " & nYKosuColGT
    With Sheets(cCfgShtName).Range("担当者状況雛型行")
        '雛型行コピー
        .Copy .Offset(nTjOff, 0)
        With .Offset(nTjOff, 0)
            '集計行罫線
            .Borders(xlEdgeTop).LineStyle = xlDouble
            '行情報セット
            .Cells(1, 1).Value = ""
            .Cells(1, 2).Value = sDateGT
            .Cells(1, 3).FormulaR1C1 = "=SUM(R[-" & nTjOff & "]C:R[-1]C)"
            .Cells(1, 4).FormulaR1C1 = "=SUM(R[-" & nTjOff & "]C:R[-1]C)"
            .Cells(1, 5).FormulaR1C1 = "=SUM(R[-" & nTjOff & "]C:R[-1]C)"
            .Cells(1, 6).FormulaR1C1 = "=SUM(R[-" & nTjOff & "]C:R[-1]C)"
            .Cells(1, 7).FormulaR1C1 = "=RC[-1]/工数１日"
            .Cells(1, 8).FormulaR1C1 = "=RC[-1]/20"
            .Cells(1, 9).FormulaR1C1 = "=(RC[-3]/RC[-6]-1)*100"
        End With
        nTjOff = nTjOff + 1
    End With
   
    '予定工数履歴の作成
    If Sheets(cCfgShtName).Range("予定工数履歴").Value = "作成する" Then
        With Sheets(cYkrShtName)
            '新規列位置の算出、雛型コピー、日付セット
            nYkrCol = .Range("予定工数履歴雛型列").Column
            Do Until .Cells(1, nYkrCol).Value = ""
                nYkrCol = nYkrCol + 1
            Loop
            .Range("予定工数履歴雛型列").Copy .Columns(nYkrCol)
            .Cells(1, nYkrCol).Value = sRDate
            
            '予定工数計、残り工数計、実績工数計のセット
            .Cells(2, nYkrCol).Value = nJKosuColGT
            .Cells(3, nYkrCol).Value = nYKosuColGT - nJKosuColGT
            .Cells(4, nYkrCol).Value = nYKosuColGT
            
            'TODO: チケット削除が発生する場合は№でのマッチング必要
            'Range("No").Resize(Range("No").Rows.Count - 3, 1).Copy .Range("No")
            'Range("題名").Resize(Range("題名").Rows.Count - 3, 1).Copy .Range("題名")
            'Range("担当者").Resize(Range("担当者").Rows.Count - 3, 1).Copy .Range("担当者")
            'Range("予定工数").Resize(Range("予定工数").Rows.Count - 3, 1).Copy .Range(.Cells(5, nYkrCol), .Cells(.Rows.Count - 1, nYkrCol))
            
            'バッファ消費率の数式セット
            .Cells(5, nYkrCol).FormulaR1C1 = "=IF(バッファ工数=0,0,ROUND((R[-1]C-R4C4)/バッファ工数*100,0))"
            
            '最終行位置の算出
            nYkrRow = 6
            Do Until .Cells(nYkrRow, 1).Value = ""
                nYkrRow = nYkrRow + 1
            Loop
            
            '予定工数セット、一致する№が無い場合は最終行以降に追加
            For i = 2 To nLastRow
                nNo = Range("No").Cells(i - 1, 1)
                Set oFind = .Range("No").Find(nNo, , xlFormulas, xlWhole)
                If Not oFind Is Nothing Then
                    .Cells(oFind.Row, nYkrCol) = Range("予定工数").Cells(i - 1, 1)
                Else
                    .Cells(nYkrRow, .Range("No").Column) = Range("No").Cells(i - 1, 1)
                    .Cells(nYkrRow, .Range("題名").Column) = Range("題名").Cells(i - 1, 1)
                    .Cells(nYkrRow, .Range("担当者").Column) = Range("担当者").Cells(i - 1, 1)
                    .Cells(nYkrRow, nYkrCol) = Range("予定工数").Cells(i - 1, 1)
                    nYkrRow = nYkrRow + 1
                End If
            Next
            
            '予定工数履歴のチケット行を№順でソート
            .Range(.Cells(6, 1), .Cells(nYkrRow, nYkrCol)).Sort _
                Key1:=.Range("No"), Order1:=xlAscending, _
                Header:=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin, _
                DataOption1:=xlSortNormal, DataOption2:=xlSortTextAsNumbers, DataOption3:=xlSortNormal
            
            'オートフィルタの再設定
            With .Range(.Columns(1), .Columns(nYkrCol))
                .AutoFilter
                .AutoFilter
            End With
        End With
    End If
    
    '進捗推移グラフの出力
    With Sheets(cYkrShtName)
        nYkrSCol = .Range("予定工数履歴雛型列").Column
        nYkrECol = nYkrSCol
        Do Until .Cells(1, nYkrECol).Value = ""
            nYkrECol = nYkrECol + 1
        Loop
        nYkrECol = nYkrECol - 1
    End With
    With Sheets(cYkrShtName)
        Set oSrcDat = .Range(.Cells(1, nYkrSCol), .Cells(4, nYkrECol))
    End With
    With Sheets(cCfgShtName)
        For Each oChtObj In .ChartObjects
             oChtObj.Delete
        Next
        Set oChtObj = .ChartObjects.Add(313, 220, 510, 300) 'Left, Top, Width, Height
    End With
    With oChtObj.Chart
        .ChartType = xlLineMarkers
        .SetSourceData Source:=oSrcDat, PlotBy:=xlRows
        With .SeriesCollection(1)
            .Name = "実績工数計"
            .Border.Weight = xlMedium
            .Border.ColorIndex = 11
            .MarkerForegroundColorIndex = 11
            .MarkerBackgroundColorIndex = 11
        End With
        With .SeriesCollection(2)
            .Name = "残り工数計"
            .Border.Weight = xlMedium
            .Border.ColorIndex = 7
            .MarkerForegroundColorIndex = 7
            .MarkerBackgroundColorIndex = 7
        End With
        With .SeriesCollection(3)
            .Name = "予定工数計"
            .Border.Weight = xlMedium
            .Border.ColorIndex = 43
            .MarkerForegroundColorIndex = 43
            .MarkerBackgroundColorIndex = 43
        End With
        .Legend.Position = xlTop
        With .Axes(xlValue)
            .Border.ColorIndex = 16
            .MajorGridlines.Border.ColorIndex = 16
        End With
        With .Axes(xlCategory)
            .CategoryType = xlCategoryScale
            .HasMajorGridlines = True
            .AxisBetweenCategories = False
            .Border.ColorIndex = 16
            .MajorGridlines.Border.ColorIndex = 16
        End With
        .PlotArea.Interior.ColorIndex = 2
    End With
    
    'バッファ管理グラフの作成
    Dim nYkrCnt, nBufKosu, nWarBufPerS, nWarBufPerE, nWarBufDif, nDanBufPerS, nDanBufPerE, nDanBufDif
    Dim arSafBufPer(), arWarBufPer(), arDanBufPer()
    
    '設定シートから設定値を取得
    With Sheets(cCfgShtName)
        nBufKosu = .Range("バッファ工数").Value
        nWarBufPerS = .Range("注意バッファ消費率開始時").Value
        nWarBufPerE = .Range("注意バッファ消費率終了時").Value
        nDanBufPerS = .Range("危険バッファ消費率開始時").Value
        nDanBufPerE = .Range("危険バッファ消費率終了時").Value
    End With
    
    'バッファ工数が設定されている場合のみ
    If nBufKosu > 0 Then
        'データ数、安全・注意・危険の％配列確保
        nYkrCnt = nYkrECol - 3
        ReDim Preserve arSafBufPer(nYkrCnt - 1)
        ReDim Preserve arWarBufPer(nYkrCnt - 1)
        ReDim Preserve arDanBufPer(nYkrCnt - 1)
        '注意・危険の増減数
        nWarBufDif = nWarBufPerE - nWarBufPerS
        nDanBufDif = nDanBufPerE - nDanBufPerS
        '予定工数履歴ごとに進捗率を算出、各％配列にセット
        With Sheets(cYkrShtName)
            For i = 0 To nYkrCnt - 1
                nShinchokuRitsu = .Cells(2, nYkrSCol + i).Value / .Cells(4, nYkrSCol + i).Value '進捗率 ＝ 実績工数 ÷ 予定工数
                arSafBufPer(i) = nWarBufPerS + Round(nWarBufDif * nShinchokuRitsu, 0)
                arWarBufPer(i) = nDanBufPerS - nWarBufPerS - Round(nWarBufDif * nShinchokuRitsu, 0) + Round(nDanBufDif * nShinchokuRitsu, 0)
                arDanBufPer(i) = 100 - nDanBufPerS - Round(nDanBufDif * nShinchokuRitsu, 0)
            Next
        End With
        
'        Debug.Print "nWarBufDif: " & nWarBufDif
'        Debug.Print "nDanBufDif: " & nDanBufDif
'        Debug.Print "arSafBufPer: " & Join(arSafBufPer, ", ")
'        Debug.Print "arWarBufPer: " & Join(arWarBufPer, ", ")
'        Debug.Print "arDanBufPer: " & Join(arDanBufPer, ", ")
        
        'バッファ管理グラフの出力
        With Sheets(cYkrShtName)
            Set oXValDat = .Range(.Cells(1, nYkrSCol), .Cells(1, nYkrECol))
            Set oBValDat = .Range(.Cells(5, nYkrSCol), .Cells(5, nYkrECol))
        End With
        With Sheets(cCfgShtName)
            Set oChtObj = .ChartObjects.Add(313, 530, 510, 300) 'Left, Top, Width, Height
        End With
        With oChtObj.Chart
            .SeriesCollection.NewSeries
            With .SeriesCollection(1)
                .ChartType = xlAreaStacked
                .XValues = oXValDat
                .Name = "安全"
                .Values = arSafBufPer
                .Interior.ColorIndex = 35 '10
                .Border.LineStyle = xlNone
            End With
            .SeriesCollection.NewSeries
            With .SeriesCollection(2)
                .Name = "注意"
                .Values = arWarBufPer
                .Interior.ColorIndex = 36 '6
                .Border.LineStyle = xlNone
            End With
            .SeriesCollection.NewSeries
            With .SeriesCollection(3)
                .Name = "危険"
                .Values = arDanBufPer
                .Interior.ColorIndex = 38 '3
                .Border.LineStyle = xlNone
            End With
            .SeriesCollection.NewSeries
            With .SeriesCollection(4)
                .ChartType = xlLineMarkers
                '.AxisGroup = xlSecondary
                .Name = "バッファ消費率"
                .Values = oBValDat
                With .Border
                    .ColorIndex = 46
                    .Weight = xlMedium
                End With
                .MarkerBackgroundColorIndex = 46
                .MarkerForegroundColorIndex = xlNone
                .MarkerStyle = xlSquare
            End With
            .Legend.Position = xlTop
            With .Axes(xlValue)
                .MaximumScale = 100
                .Border.ColorIndex = 16
                .MajorGridlines.Border.ColorIndex = 16
            End With
            With .Axes(xlCategory)
                .CategoryType = xlCategoryScale
                .HasMajorGridlines = True
                .AxisBetweenCategories = False
                .Border.ColorIndex = 16
                .MajorGridlines.Border.ColorIndex = 16
            End With
        End With
    End If

End Sub

'営業日で日数加算
Private Function WorkDateAdd(sDate, nDay)
    'WorkDateAdd = Format(CDate(sDate) + nDay, "yyyy/mm/dd")
    'Exit Function
    
    '指定した日が営業日かどうかもチェック
    nAdd = -1
    nDate = CDate(sDate) - 1
    Do
        nDate = nDate + 1
        bWorkDay = True
        
        '休日の曜日判定
        If InStr(sHolidayOfWeek, Weekday(nDate)) > 0 Then
            bWorkDay = False
        End If
        
        '祝日判定
        Dim oFind
        Set oFind = Sheets(cCfgShtName).Range("祝日設定").Find(nDate, , xlFormulas, xlWhole)
        If Not oFind Is Nothing Then
            bWorkDay = False
        End If
        
        '休日祝日の無視判定
        If Not bWorkDay Then
            Set oFind = Sheets(cCfgShtName).Range("無視休日祝日").Find(nDate, , xlFormulas, xlWhole)
            If Not oFind Is Nothing Then
                bWorkDay = True
            End If
        End If
        
        '営業日ならカウント
        If bWorkDay Then
            nAdd = nAdd + 1
        End If
    Loop While nAdd < nDay
    WorkDateAdd = Format(nDate, "yyyy/mm/dd")
End Function

