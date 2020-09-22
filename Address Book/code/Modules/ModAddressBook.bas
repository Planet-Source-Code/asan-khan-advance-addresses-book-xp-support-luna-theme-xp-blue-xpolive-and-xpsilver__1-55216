Attribute VB_Name = "ModAddressBook"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                        (ByVal hWnd As Long, ByVal lpOperation As String, _
                        ByVal lpFile As String, ByVal lpParameters As String, _
                        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public CN               As ADODB.Connection
Public StrSQL           As String
Private MyForm          As Frm_addressbook

Sub Main()

    Set CN = New ADODB.Connection
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\addressbook.mdb"
    ' check connection status
    If CN.State Then
        ' connection OK
        Set MyForm = New Frm_addressbook
        MyForm.Show

      Else
        MsgBoxXP "Could not connect to database ...", vbExclamation, "Connection failed", , , 3
        Set CN = Nothing
    End If

End Sub

Public Sub EndProgram()

    CN.Close
    Set CN = Nothing
    Set MyForm = Nothing

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2004-Jul-28 14:07) 10 + 27 = 37 Lines
