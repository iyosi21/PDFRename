Attribute VB_Name = "Module1"
Sub renamex()
    '保存先
    MsgBox "PDFファイルをリネームします。B列が空欄の場合はリネームされません。"
    myfolder$ = ThisWorkbook.PATH & "\"
    Const A& = 1, B& = 2
    Dim YorN As Integer
    Dim SecPATH As String: SecPATH = Worksheets(1).Cells(1, 8).Value
    
    If SecPATH <> myfolder Then  '1
        YorN = MsgBox("下記のパスのPDFをリネームをしますか？" & Chr(13) & _
                        SecPATH, vbYesNo)
        If YorN = vbYes Then '2
            myfolder = SecPATH
        Else '2
            MsgBox "このブックのPDFをリネームします。"
            If Dir(myfolder & Worksheets(1).Cells(1, 1).Value) <> "" Then '3
            Else '3
                MsgBox "ファイルが見つかりません", vbExclamation
                Exit Sub
            End If '3
        End If '2
    End If '1

    Dim con As Long
    con = Cells(Rows.Count, A).End(xlUp).Row
    Dim i As Long
    For i = 1 To con
        If Cells(i, 2).Value <> "" Then
          On Error GoTo MyError
          Name myfolder & Cells(i, A).Value As myfolder & Cells(i, B).Value
        End If
110
    Next i
    MsgBox "PDFをリネームしました"
    Exit Sub
    
MyError:
      MsgBox "フォルダに同じ名前のファイルがあります。B列：" & i & " 番目", vbExclamation
      GoTo 110
    
End Sub

Sub Get_PDFFile()
MsgBox "フォルダ内のPDFファイル名を取得します。"
Worksheets(1).Columns(1).ClearContents

Dim PATH As String: PATH = ThisWorkbook.PATH & "\"
Dim str As String
str = Dir(PATH & "*.pdf")

Worksheets(1).Cells(1, 8).Value = ThisWorkbook.PATH & "\"

i = 0
Do While str <> ""
    i = i + 1
    ActiveSheet.Cells(i, 1).Value = str
    str = Dir()
Loop

MsgBox "PDFのファイル名を取得しました"
End Sub

Sub Get_Folder()
Worksheets(1).Cells(1, 8).Value = ""

Dim Shell, myPath
Dim str As String
Dim i As Integer

Set Shell = CreateObject("shell.application")
Set myPath = Shell.browseforfolder(&O0, "フォルダを選んでください", &H1 + &H10, "desktop")
If Not myPath Is Nothing Then
    GoTo 100
Else
    MsgBox "フォルダが見つかりません", vbCritical
    Exit Sub
End If
Worksheets(1).Columns(1).ClearContents

100 'フォルダが見つかるとここに飛んでくる。

Worksheets(1).Cells(1, 8).Value = myPath.self.PATH & "\"

str = Dir(myPath.self.PATH & "\" & "*.pdf*")

i = 0
Do While str <> ""
    i = i + 1
    ActiveSheet.Cells(i, 1).Value = str
    str = Dir()
Loop

End Sub
