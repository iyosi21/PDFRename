Attribute VB_Name = "Module1"
Sub renamex()
    '�ۑ���
    MsgBox "PDF�t�@�C�������l�[�����܂��BB�񂪋󗓂̏ꍇ�̓��l�[������܂���B"
    myfolder$ = ThisWorkbook.PATH & "\"
    Const A& = 1, B& = 2
    Dim YorN As Integer
    Dim SecPATH As String: SecPATH = Worksheets(1).Cells(1, 8).Value
    
    If SecPATH <> myfolder Then  '1
        YorN = MsgBox("���L�̃p�X��PDF�����l�[�������܂����H" & Chr(13) & _
                        SecPATH, vbYesNo)
        If YorN = vbYes Then '2
            myfolder = SecPATH
        Else '2
            MsgBox "���̃u�b�N��PDF�����l�[�����܂��B"
            If Dir(myfolder & Worksheets(1).Cells(1, 1).Value) <> "" Then '3
            Else '3
                MsgBox "�t�@�C����������܂���", vbExclamation
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
    MsgBox "PDF�����l�[�����܂���"
    Exit Sub
    
MyError:
      MsgBox "�t�H���_�ɓ������O�̃t�@�C��������܂��BB��F" & i & " �Ԗ�", vbExclamation
      GoTo 110
    
End Sub

Sub Get_PDFFile()
MsgBox "�t�H���_����PDF�t�@�C�������擾���܂��B"
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

MsgBox "PDF�̃t�@�C�������擾���܂���"
End Sub

Sub Get_Folder()
Worksheets(1).Cells(1, 8).Value = ""

Dim Shell, myPath
Dim str As String
Dim i As Integer

Set Shell = CreateObject("shell.application")
Set myPath = Shell.browseforfolder(&O0, "�t�H���_��I��ł�������", &H1 + &H10, "desktop")
If Not myPath Is Nothing Then
    GoTo 100
Else
    MsgBox "�t�H���_��������܂���", vbCritical
    Exit Sub
End If
Worksheets(1).Columns(1).ClearContents

100 '�t�H���_��������Ƃ����ɔ��ł���B

Worksheets(1).Cells(1, 8).Value = myPath.self.PATH & "\"

str = Dir(myPath.self.PATH & "\" & "*.pdf*")

i = 0
Do While str <> ""
    i = i + 1
    ActiveSheet.Cells(i, 1).Value = str
    str = Dir()
Loop

End Sub
