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

    '<�萔>------------------------------------------------------------------------------------------
    
    Const int_indent As Integer = 4 '������
    Const str_nlineCode As String = vbCrLf '���s�R�[�h
    '
    '-----------------------------------------------------------------------------------------</�萔>

    '�ϐ��錾
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
    
    '�V�[�g�͈͑S�I������Ă����ꍇ�́AUsedRange���Ɏ��܂�悤�Ƀg���~���O
    Set range_selection = trimWithUsedRange(Selection)
    
    '������
    ReDim strarr_builder(0 To 0)
    
    '<��\���s�A��\����̉��>-----------------------
    
    lng_startOfRowPre = range_selection.Row
    lng_lastOfRowPre = lng_startOfRow + range_selection.Rows.Count - 1
    lng_startOfColPre = range_selection.Column
    lng_lastOfColPre = lng_startOfCol + range_selection.Columns.Count - 1
    
    '��\���s�ԍ��̃R���N�V�������쐬
    For lng_rowIdx = lng_startOfRowPre To lng_lastOfRowPre
        If Application.ActiveSheet.Rows(lng_rowIdx).Hidden Then ' ��\���s�̏ꍇ
            clctn_hiddenRows.Add Item:=lng_rowIdx, Key:=(CStr(lng_rowIdx))
        End If
    Next
    
    '��\����ԍ��̃R���N�V�������쐬
    For lng_colIdx = lng_startOfColPre To lng_lastOfColPre
        If Application.ActiveSheet.Columns(lng_colIdx).Hidden Then ' ��\����̏ꍇ
            clctn_hiddenCols.Add Item:=lng_colIdx, Key:=CStr(lng_colIdx)
        End If
    Next
    
    
    '----------------------</��\���s�A��\����̉��>
    
    '<table> �^�O
    strarr_builder(UBound(strarr_builder)) = func_getTagStart("table")
    
    '�\������Ă��邱��(��\���ł͂Ȃ�����)
    'mergearea �������Ă���ꍇ�́A��\���̍s�A��A�R���N�V�����ԍ�����������
    '������������ŁA�������Ă���Z���ʒu�����ゾ�����ꍇ�A������������Ԃ̍s����rowspan=""�A�񐔂�colspan="" �ɔ��f����
    
    For lng_rowIdx = lng_startOfRowPre To lng_lastOfRowPre
    
        If Not func_isInCollection(CStr(lng_rowIdx), clctn_hiddenRows) Then ' ��\���s�ł͂Ȃ��ꍇ
            
            Dim strarr_builder_line() As String
            ReDim strarr_builder_line(0 To 0)
            isFirstCol = True
    
            For lng_colIdx = lng_startOfColPre To lng_lastOfColPre
                
                If Not func_isInCollection(CStr(lng_colIdx), clctn_hiddenCols) Then ' ��\����ł͂Ȃ��ꍇ
                    
                    Set rng_cellOfInterest = Application.ActiveSheet.Cells(lng_rowIdx, lng_colIdx)
                    
                    If Not rng_cellOfInterest.MergeCells Then ' �����Z���ł͂Ȃ��ꍇ
                        
                        ' �\������Ă���P��Z���Ƃ݂Ȃ�(= <td> �v�f�ƂȂ�Z���Ƃ݂Ȃ�)
                        If isFirstCol Then ' <tr> �v�f���̍ŏ��� <td> �ƂȂ�Ȃ�
                            isFirstCol = False
                        Else ' <tr> �v�f���̍ŏ��� <td> �ƂȂ�Ȃ�
                            ReDim Preserve strarr_builder_line(0 To UBound(strarr_builder_line) + 1)
                        End If
                        strarr_builder_line(UBound(strarr_builder_line)) = func_getTagStart("td") & rng_cellOfInterest.Text & func_getTagEnd("td") '<td> �v�f�̒�`
                        
                    Else '�����Z���̏ꍇ
                        
                        ' �\������Ă��錋���Z���Ƃ݂Ȃ�
                        
                        '�s�ԍ��R���N�V�����̍쐬
                        Set clctn_mergedRowNums = Nothing
                        lng_mergeRowStart = rng_cellOfInterest.MergeArea.Cells(1, 1).Row '�����Z���̍s�J�n�ԍ�
                        lng_mergeRowEnd = lng_mergeRowStart + rng_cellOfInterest.MergeArea.Rows.Count - 1 '�����Z���̍s�I���ԍ�
                        For lng_rowCollector = lng_mergeRowStart To lng_mergeRowEnd
                            If func_isInCollection(CStr(lng_rowCollector), clctn_hiddenRows) Then ' ��\���s�ł͂Ȃ��ꍇ
                                clctn_mergedRowNums.Add Item:=lng_rowCollector, Key:=CStr(lng_rowCollector) '�s�ԍ��R���N�V�����ɒǉ�
                            End If
                        Next
                        
                        '��ԍ��R���N�V�����̍쐬
                        Set clctn_mergedColNums = Nothing
                        lng_mergeColStart = rng_cellOfInterest.MergeArea.Cells(1, 1).Column '�����Z���̗�J�n�ԍ�
                        lng_mergeColEnd = lng_mergeColStart + rng_cellOfInterest.MergeArea.Rows.Count - 1 '�����Z���̗�I���ԍ�
                        For lng_colCollector = lng_mergeColStart To lng_mergeColEnd
                            If func_isInCollection(CStr(lng_colCollector), clctn_hiddenCols) Then ' ��\����ł͂Ȃ��ꍇ
                                clctn_mergedColNums.Add Item:=lng_colCollector, Key:=CStr(lng_colCollector) '��ԍ��R���N�V�����ɒǉ�
                            End If
                        Next
                        
                        ' �������Ă���Z���ʒu�����ォ�ǂ���
                        If (rng_cellOfInterest.Cells(1, 1).Row = clctn_mergedRowNums(1)) Then '�������Ă���Z�����ł���̏ꍇ
                            If (rng_cellOfInterest.Cells(1, 1).Column = clctn_mergedColNums(1)) Then '�������Ă���Z�����ł����̏ꍇ
                            
                                Dim dict_tag_properties As Object
                                Set dict_tag_properties = Nothing
                                
                                If (1 < clctn_mergedRowNums.Count) Or (1 < clctn_mergedColNums.Count) Then ' �����Z���� 2 �ȏ゠��ꍇ
                                    
                                    Set dict_tag_properties = CreateObject("Scripting.Dictionary")
                                    
                                    If 1 < clctn_mergedRowNums.Count Then ' �������ꂽ�s�Z����2�ȏ゠��ꍇ
                                        dict_tag_properties.Add Key:="rowspan", Item:=CStr(clctn_mergedRowNums.Count)
                                    End If
                                    
                                    If 1 < clctn_mergedColNums.Count Then ' �������ꂽ��Z����2�ȏ゠��ꍇ
                                        dict_tag_properties.Add Key:="colspan", Item:=CStr(clctn_mergedColNums.Count)
                                    End If
                                
                                End If
                                
                                ' <td> �v�f�Ƃ݂Ȃ�
                                If isFirstCol Then ' <tr> �v�f���̍ŏ��� <td> �ƂȂ�Ȃ�
                                    isFirstCol = False
                                Else ' <tr> �v�f���̍ŏ��� <td> �ƂȂ�Ȃ�
                                    ReDim Preserve strarr_builder_line(0 To UBound(strarr_builder_line) + 1)
                                End If
                                
                                strarr_builder_line(UBound(strarr_builder_line)) = func_getTagStart("td", dict_tag_properties) & rng_cellOfInterest.Text & func_getTagEnd("td") '<td> �v�f�̒�`
                                
                            End If
                        End If
                        
                    End If
                    
                End If
            
            ReDim Preserve strarr_builder(0 To UBound(strarr_builder) + 1)
            strarr_builder(UBound(strarr_builder)) = func_getTagStart("tr") & Join(strarr_builder_line, "") & func_getTagEnd("tr")
            
            Next
        
        End If
        
    Next
    
    
    '</table> �^�O
    ReDim Preserve strarr_builder(0 To UBound(strarr_builder) + 1)
    strarr_builder(UBound(strarr_builder)) = func_getTagEnd("table")
    
    '�z��̕�����
    str_tmp = Join(strarr_builder, str_nlineCode)
    
    Debug.Print str_tmp
    SetCB str_tmp
    
End Sub

Private Function func_getTagStart(ByVal str_tagname, Optional ByRef strclctn_properties As Collection = Nothing) As String
    Dim str_properties As String
    If strclctn_properties Is Nothing Then ' �v���p�e�B���w�肳��Ă���ꍇ
        
    Else
        str_properties = ""
    End If
    
    func_getTagStart = "<" + str_tagname + ">"
End Function

Private Function func_getTagEnd(ByVal str_tagname) As String
    func_getTagEnd = "<" + "/" + str_tagname + ">"
End Function

'
' �R���N�V�������Ɏw��L�[�����݂��邩�ǂ�����Ԃ�
'
Private Function func_isInCollection(ByVal str_item, ByRef clctn As Collection) As Boolean
    Dim bl_ret As Boolean
    
    bl_ret = False
    
On Error GoTo NONE: ' �G���[������������ Catch �ֈړ�����
    clctn.Item (str_item)
    bl_ret = True '�G���[���������Ȃ��ꍇ�́ATrue ��������
    
NONE:
    'nothing to do
    
    func_isInCollection = bl_ret
    
End Function


'
' �Z���Q�Ɣ͈͂� UsedRange �͈͂Ɏ��܂�悤�Ƀg���~���O����
'
Private Function trimWithUsedRange(ByVal rangeObj As Range) As Range

    'variables
    Dim ret As Range
    Dim long_bottom_right_row_idx_of_specified As Long
    Dim long_bottom_right_col_idx_of_specified As Long
    Dim long_bottom_right_row_idx_of_used As Long
    Dim long_bottom_right_col_idx_of_used As Long

    '�w��͈͂̉E���ʒu�̎擾
    long_bottom_right_row_idx_of_specified = rangeObj.Item(1).Row + rangeObj.Rows.Count - 1
    long_bottom_right_col_idx_of_specified = rangeObj.Item(1).Column + rangeObj.Columns.Count - 1
    
    'UsedRange�̉E���ʒu�̎擾
    With rangeObj.Parent.UsedRange
        long_bottom_right_row_idx_of_used = .Item(1).Row + .Rows.Count - 1
        long_bottom_right_col_idx_of_used = .Item(1).Column + .Columns.Count - 1
    End With
    
    '�g���~���O
    Set ret = rangeObj.Parent.Range( _
        rangeObj.Item(1), _
        rangeObj.Parent.Cells( _
            IIf(long_bottom_right_row_idx_of_specified > long_bottom_right_row_idx_of_used, long_bottom_right_row_idx_of_used, long_bottom_right_row_idx_of_specified), _
            IIf(long_bottom_right_col_idx_of_specified > long_bottom_right_col_idx_of_used, long_bottom_right_col_idx_of_used, long_bottom_right_col_idx_of_specified) _
        ) _
    )
    
    '�i�[���ďI��
    Set trimWithUsedRange = ret
    
End Function


'<�N���b�v�{�[�h����>-------------------------------------------

'�N���b�v�{�[�h�ɕ�������i�[
Private Sub SetCB(ByVal str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .Text = str
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub

'�N���b�v�{�[�h���當������擾
Private Sub GetCB(ByRef str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    If .CanPaste = True Then .Paste
    str = .Text
  End With
End Sub

'------------------------------------------</�N���b�v�{�[�h����>

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




