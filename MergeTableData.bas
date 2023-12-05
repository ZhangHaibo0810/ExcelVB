Attribute VB_Name = "模块1"
Sub MergeTableData()

'定义合并文件夹目录
Dim path As String

'定义合并总表的文件名
Dim activeName As String

'定义当前文件夹当前检查获取的文件名
Dim xlsxName As String

'定义wb存储获取的工作簿
Dim Wb As Workbook

'关闭屏幕更新，优化合并效率
Application.ScreenUpdating = False

Dim alldic, dicIdx

'获取当前合并总表的目录, 'E:\多表本合并'
path = ActiveWorkbook.path

'选择文件夹
With Application.FileDialog(msoFileDialogFolderPicker)
    If .Show Then
        path = .SelectedItems(1)
    End If
End With

'获取目前文件夹下所有目录
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

'获取当前合并后总表的文件名
activeName = ActiveWorkbook.Name

For Each Key In alldic.keys
    '获取path路径下的所有'.xlsx'文件名，'E:\多表本合并\*.xlsx'
    xlsxName = Dir(Key & "\" & "*.xlsx")
    
    '当前文件夹内的xlsx文件未遍历完
    Do While xlsxName <> ""
        '并且当前访问的不是总表
        If xlsxName <> activeName Then
            '依次打开每一个xlsx文件
            Set Wb = Workbooks.Open(Key & "\" & xlsxName)
            
            ToMergeRows = Wb.Sheets(1).UsedRange.Rows.Count
            
            UsedRows = ThisWorkbook.Worksheets(1).UsedRange.Rows.Count
    
            Wb.Sheets(1).Range(Cells(5, 2), Cells(ToMergeRows + 1, 2)).Copy ThisWorkbook.Worksheets(1).Cells(UsedRows + 1, 2)
            Wb.Sheets(1).Range(Cells(5, 4), Cells(ToMergeRows + 1, 4)).Copy ThisWorkbook.Worksheets(1).Cells(UsedRows + 1, 4)
            Wb.Sheets(1).Range(Cells(5, 6), Cells(ToMergeRows + 1, 6)).Copy ThisWorkbook.Worksheets(1).Cells(UsedRows + 1, 6)
            Wb.Sheets(1).Range(Cells(5, 12), Cells(ToMergeRows + 1, 140)).Copy ThisWorkbook.Worksheets(1).Cells(UsedRows + 1, 12)
            
            '关闭当前遍历的工作簿，不保存
            Wb.Close False
    
        End If
    
        '调用Dir函数，找到当前目录的下一个xlsx文件
        xlsxName = Dir
    
    Loop

Next
'恢复幕更新
Application.ScreenUpdating = True


End Sub
