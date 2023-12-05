Attribute VB_Name = "ģ��1"
Sub MergeTableData()

'����ϲ��ļ���Ŀ¼
Dim path As String

'����ϲ��ܱ���ļ���
Dim activeName As String

'���嵱ǰ�ļ��е�ǰ����ȡ���ļ���
Dim xlsxName As String

'����wb�洢��ȡ�Ĺ�����
Dim Wb As Workbook

'�ر���Ļ���£��Ż��ϲ�Ч��
Application.ScreenUpdating = False

Dim alldic, dicIdx

'��ȡ��ǰ�ϲ��ܱ��Ŀ¼, 'E:\����ϲ�'
path = ActiveWorkbook.path

'ѡ���ļ���
With Application.FileDialog(msoFileDialogFolderPicker)
    If .Show Then
        path = .SelectedItems(1)
    End If
End With

'��ȡĿǰ�ļ���������Ŀ¼
Set alldic = CreateObject("Scripting.Dictionary")

alldic.Add path, ""

dicIdx = 0
Do While dicIdx < alldic.Count
    Key = alldic.keys
    newdic = Dir(Key(dicIdx) & "\", vbDirectory)
    Do While newdic <> ""
        If (newdic <> ".") And (newdic <> "..") Then
            If (GetAttr(Key(dicIdx) & "\" & newdic) And vbDirectory) = vbDirectory Then
                alldic.Add Key(dicIdx) & "\" & newdic, ""
            End If
        End If
        newdic = Dir()
    Loop
    dicIdx = dicIdx + 1
Loop

'��ȡ��ǰ�ϲ����ܱ���ļ���
activeName = ActiveWorkbook.Name

For Each Key In alldic.keys
    '��ȡpath·���µ�����'.xlsx'�ļ�����'E:\����ϲ�\*.xlsx'
    xlsxName = Dir(Key & "\" & "*.xlsx")
    
    '��ǰ�ļ����ڵ�xlsx�ļ�δ������
    Do While xlsxName <> ""
        '���ҵ�ǰ���ʵĲ����ܱ�
        If xlsxName <> activeName Then
            '���δ�ÿһ��xlsx�ļ�
            Set Wb = Workbooks.Open(Key & "\" & xlsxName)
            
            ToMergeRows = Wb.Sheets(1).UsedRange.Rows.Count
            
            UsedRows = ThisWorkbook.Worksheets(1).UsedRange.Rows.Count
    
            Wb.Sheets(1).Range(Cells(5, 2), Cells(ToMergeRows + 1, 2)).Copy ThisWorkbook.Worksheets(1).Cells(UsedRows + 1, 2)
            Wb.Sheets(1).Range(Cells(5, 4), Cells(ToMergeRows + 1, 4)).Copy ThisWorkbook.Worksheets(1).Cells(UsedRows + 1, 4)
            Wb.Sheets(1).Range(Cells(5, 6), Cells(ToMergeRows + 1, 6)).Copy ThisWorkbook.Worksheets(1).Cells(UsedRows + 1, 6)
            Wb.Sheets(1).Range(Cells(5, 12), Cells(ToMergeRows + 1, 140)).Copy ThisWorkbook.Worksheets(1).Cells(UsedRows + 1, 12)
            
            '�رյ�ǰ�����Ĺ�������������
            Wb.Close False
    
        End If
    
        '����Dir�������ҵ���ǰĿ¼����һ��xlsx�ļ�
        xlsxName = Dir
    
    Loop

Next
'�ָ�Ļ����
Application.ScreenUpdating = True


End Sub
