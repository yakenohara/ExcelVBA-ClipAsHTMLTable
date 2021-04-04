Attribute VB_Name = "Module1"
'<License>------------------------------------------------------------
'
' Copyright (c) 2021 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>


Sub test()

    '<定数>------------------------------------------------------------------------------------------
    
    Const int_indent As Integer = 4 '字下げ
    Const str_nlineCode As String = vbCrLf '改行コード
    '
    '-----------------------------------------------------------------------------------------</定数>

    '変数宣言
    Dim strarr_builder() As String
    
    Dim clctn_hiddenRows As New Collection
    Dim clctn_hiddenCols As New Collection
    Dim clctn_mergedRowNums As New Collection
    Dim clctn_mergedColNums As New Collection
    
    Dim lng_startOfRowPre As Long
    Dim lng_lastOfRowPre As Long
    Dim lng_startOfColPre As Long
    Dim lng_lastOfColPre As Long
    
    Dim lng_startOfRow As Long
    Dim lng_lastOfRow As Long
    Dim lng_startOfCol As Long
    Dim lng_lastOfCol As Long
    
    Dim isFirstCol As Boolean
    
    'シート範囲全選択されていた場合は、UsedRange内に収まるようにトリミング
    Set range_selection = trimWithUsedRange(Selection)
    
    '初期化
    ReDim strarr_builder(0 To 0)
    
    '<非表示行、非表示列の解析>-----------------------
    
    lng_startOfRowPre = range_selection.Row
    lng_lastOfRowPre = lng_startOfRow + range_selection.Rows.Count - 1
    lng_startOfColPre = range_selection.Column
    lng_lastOfColPre = lng_startOfCol + range_selection.Columns.Count - 1
    
    '非表示行番号のコレクションを作成
    For lng_rowIdx = lng_startOfRowPre To lng_lastOfRowPre
        If Application.ActiveSheet.Rows(lng_rowIdx).Hidden Then ' 非表示行の場合
            clctn_hiddenRows.Add Item:=lng_rowIdx, Key:=(CStr(lng_rowIdx))
        End If
    Next
    
    '非表示列番号のコレクションを作成
    For lng_colIdx = lng_startOfColPre To lng_lastOfColPre
        If Application.ActiveSheet.Columns(lng_colIdx).Hidden Then ' 非表示列の場合
            clctn_hiddenCols.Add Item:=lng_colIdx, Key:=CStr(lng_colIdx)
        End If
    Next
    
    
    '----------------------</非表示行、非表示列の解析>
    
    '<table> タグ
    strarr_builder(UBound(strarr_builder)) = func_getTagStart("table")
    
    '表示されていること(非表示ではないこと)
    'mergearea を持っている場合は、非表示の行、列、コレクション番号を差っ引く
    '差っ引いた上で、注視しているセル位置が左上だった場合、差っ引いた状態の行数をrowspan=""、列数をcolspan="" に反映する
    
    For lng_rowIdx = lng_startOfRowPre To lng_lastOfRowPre
    
        If Not func_isInCollection(CStr(lng_rowIdx), clctn_hiddenRows) Then ' 非表示行ではない場合
            
            Dim strarr_builder_line() As String
            ReDim strarr_builder_line(0 To 0)
            isFirstCol = True
    
            For lng_colIdx = lng_startOfColPre To lng_lastOfColPre
                
                If Not func_isInCollection(CStr(lng_colIdx), clctn_hiddenCols) Then ' 非表示列ではない場合
                    
                    Set rng_cellOfInterest = Application.ActiveSheet.Cells(lng_rowIdx, lng_colIdx)
                    
                    If Not rng_cellOfInterest.MergeCells Then ' 結合セルではない場合
                        
                        ' 表示されている単一セルとみなす(= <td> 要素となるセルとみなす)
                        If isFirstCol Then ' <tr> 要素内の最初の <td> となるなら
                            isFirstCol = False
                        Else ' <tr> 要素内の最初の <td> となるなら
                            ReDim Preserve strarr_builder_line(0 To UBound(strarr_builder_line) + 1)
                        End If
                        strarr_builder_line(UBound(strarr_builder_line)) = func_getTagStart("td") & rng_cellOfInterest.Text & func_getTagEnd("td") '<td> 要素の定義
                        
                    Else '結合セルの場合
                        
                        ' 表示されている結合セルとみなす
                        
                        '行番号コレクションの作成
                        Set clctn_mergedRowNums = Nothing
                        lng_mergeRowStart = rng_cellOfInterest.MergeArea.Cells(1, 1).Row '結合セルの行開始番号
                        lng_mergeRowEnd = lng_mergeRowStart + rng_cellOfInterest.MergeArea.Rows.Count - 1 '結合セルの行終了番号
                        For lng_rowCollector = lng_mergeRowStart To lng_mergeRowEnd
                            If func_isInCollection(CStr(lng_rowCollector), clctn_hiddenRows) Then ' 非表示行ではない場合
                                clctn_mergedRowNums.Add Item:=lng_rowCollector, Key:=CStr(lng_rowCollector) '行番号コレクションに追加
                            End If
                        Next
                        
                        '列番号コレクションの作成
                        Set clctn_mergedColNums = Nothing
                        lng_mergeColStart = rng_cellOfInterest.MergeArea.Cells(1, 1).Column '結合セルの列開始番号
                        lng_mergeColEnd = lng_mergeColStart + rng_cellOfInterest.MergeArea.Rows.Count - 1 '結合セルの列終了番号
                        For lng_colCollector = lng_mergeColStart To lng_mergeColEnd
                            If func_isInCollection(CStr(lng_colCollector), clctn_hiddenCols) Then ' 非表示列ではない場合
                                clctn_mergedColNums.Add Item:=lng_colCollector, Key:=CStr(lng_colCollector) '列番号コレクションに追加
                            End If
                        Next
                        
                        ' 注視しているセル位置が左上かどうか
                        If (rng_cellOfInterest.Cells(1, 1).Row = clctn_mergedRowNums(1)) Then '注視しているセルが最も上の場合
                            If (rng_cellOfInterest.Cells(1, 1).Column = clctn_mergedColNums(1)) Then '注視しているセルが最も左の場合
                            
                                Dim dict_tag_properties As Object
                                Set dict_tag_properties = Nothing
                                
                                If (1 < clctn_mergedRowNums.Count) Or (1 < clctn_mergedColNums.Count) Then ' 結合セルが 2 つ以上ある場合
                                    
                                    Set dict_tag_properties = CreateObject("Scripting.Dictionary")
                                    
                                    If 1 < clctn_mergedRowNums.Count Then ' 結合された行セルが2つ以上ある場合
                                        dict_tag_properties.Add Key:="rowspan", Item:=CStr(clctn_mergedRowNums.Count)
                                    End If
                                    
                                    If 1 < clctn_mergedColNums.Count Then ' 結合された列セルが2つ以上ある場合
                                        dict_tag_properties.Add Key:="colspan", Item:=CStr(clctn_mergedColNums.Count)
                                    End If
                                
                                End If
                                
                                ' <td> 要素とみなす
                                If isFirstCol Then ' <tr> 要素内の最初の <td> となるなら
                                    isFirstCol = False
                                Else ' <tr> 要素内の最初の <td> となるなら
                                    ReDim Preserve strarr_builder_line(0 To UBound(strarr_builder_line) + 1)
                                End If
                                
                                strarr_builder_line(UBound(strarr_builder_line)) = func_getTagStart("td", dict_tag_properties) & rng_cellOfInterest.Text & func_getTagEnd("td") '<td> 要素の定義
                                
                            End If
                        End If
                        
                    End If
                    
                End If
            
            ReDim Preserve strarr_builder(0 To UBound(strarr_builder) + 1)
            strarr_builder(UBound(strarr_builder)) = func_getTagStart("tr") & Join(strarr_builder_line, "") & func_getTagEnd("tr")
            
            Next
        
        End If
        
    Next
    
    
    '</table> タグ
    ReDim Preserve strarr_builder(0 To UBound(strarr_builder) + 1)
    strarr_builder(UBound(strarr_builder)) = func_getTagEnd("table")
    
    '配列の文字列化
    str_tmp = Join(strarr_builder, str_nlineCode)
    
    Debug.Print str_tmp
    SetCB str_tmp
    
End Sub

Private Function func_getTagStart(ByVal str_tagname, Optional ByRef strclctn_properties As Collection = Nothing) As String
    Dim str_properties As String
    If strclctn_properties Is Nothing Then ' プロパティが指定されている場合
        
    Else
        str_properties = ""
    End If
    
    func_getTagStart = "<" + str_tagname + ">"
End Function

Private Function func_getTagEnd(ByVal str_tagname) As String
    func_getTagEnd = "<" + "/" + str_tagname + ">"
End Function

'
' コレクション内に指定キーが存在するかどうかを返す
'
Private Function func_isInCollection(ByVal str_item, ByRef clctn As Collection) As Boolean
    Dim bl_ret As Boolean
    
    bl_ret = False
    
On Error GoTo NONE: ' エラーが発生したら Catch へ移動する
    clctn.Item (str_item)
    bl_ret = True 'エラーが発生しない場合は、True を代入する
    
NONE:
    'nothing to do
    
    func_isInCollection = bl_ret
    
End Function


'
' セル参照範囲が UsedRange 範囲に収まるようにトリミングする
'
Private Function trimWithUsedRange(ByVal rangeObj As Range) As Range

    'variables
    Dim ret As Range
    Dim long_bottom_right_row_idx_of_specified As Long
    Dim long_bottom_right_col_idx_of_specified As Long
    Dim long_bottom_right_row_idx_of_used As Long
    Dim long_bottom_right_col_idx_of_used As Long

    '指定範囲の右下位置の取得
    long_bottom_right_row_idx_of_specified = rangeObj.Item(1).Row + rangeObj.Rows.Count - 1
    long_bottom_right_col_idx_of_specified = rangeObj.Item(1).Column + rangeObj.Columns.Count - 1
    
    'UsedRangeの右下位置の取得
    With rangeObj.Parent.UsedRange
        long_bottom_right_row_idx_of_used = .Item(1).Row + .Rows.Count - 1
        long_bottom_right_col_idx_of_used = .Item(1).Column + .Columns.Count - 1
    End With
    
    'トリミング
    Set ret = rangeObj.Parent.Range( _
        rangeObj.Item(1), _
        rangeObj.Parent.Cells( _
            IIf(long_bottom_right_row_idx_of_specified > long_bottom_right_row_idx_of_used, long_bottom_right_row_idx_of_used, long_bottom_right_row_idx_of_specified), _
            IIf(long_bottom_right_col_idx_of_specified > long_bottom_right_col_idx_of_used, long_bottom_right_col_idx_of_used, long_bottom_right_col_idx_of_specified) _
        ) _
    )
    
    '格納して終了
    Set trimWithUsedRange = ret
    
End Function


'<クリップボード操作>-------------------------------------------

'クリップボードに文字列を格納
Private Sub SetCB(ByVal str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .Text = str
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub

'クリップボードから文字列を取得
Private Sub GetCB(ByRef str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    If .CanPaste = True Then .Paste
    str = .Text
  End With
End Sub

'------------------------------------------</クリップボード操作>

Sub aaa()
'    Dim clctn_hiddenRows As New Collection
'    clctn_hiddenRows.Add Item:=2, Key:=CStr(2)
'    'Debug.Print clctn_hiddenRows.Count
'    For Each kkk In clctn_hiddenRows.Keys
'        Debug.Print kkk
'    Next
'    Debug.Print func_isInCollection("2", clctn_hiddenRows)
'    Debug.Print Application.ActiveSheet.Cells(4, 4).MergeArea.Columns.Count
'    Debug.Print Application.ActiveSheet.Cells(4, 4).MergeArea.Row

    Dim dict_view_setting As Object
    Set dict_view_setting = CreateObject("Scripting.Dictionary")

    With dict_view_setting
        .Add "prop_int_zoom_level", "int_tmp_val"
        .Add "prop_bool_top_left_option_enabled", "SetSameViewFormMod.chkbx_top_left.Value"
        .Add "prop_str_top_left_address_of_view", "SetSameViewFormMod.txtbx_top_left_address_of_view.Value"
        .Add "prop_str_range_address_to_select", "SetSameViewFormMod.txtbx_range_address_to_select.Value"
        .Add "prop_str_sheet_name_to_activate", "SetSameViewFormMod.cmbbx_sheet_name_to_activate.Value"
        .Add "prop_bool_minimize_ribbon_option_enabled", "SetSameViewFormMod.chkbx_minimize_ribbon.Value"
        .Add "prop_bool_all_books_option_enabled", "SetSameViewFormMod.chkbx_all_books.Value"
    End With
    x = bbb(dict_view_setting)
End Sub

Private Function bbb(Optional ByVal dict_view_setting As Object = Nothing)
    For Each vvv In dict_view_setting
        Debug.Print "key:" & vvv & ", val:" & dict_view_setting.Item(vvv)
    Next vvv
End Function




