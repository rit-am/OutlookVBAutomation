VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents Items As Outlook.Items
Attribute Items.VB_VarHelpID = -1
Private Sub Application_Startup()
  Dim olApp As Outlook.Application
  Dim objNS As Outlook.NameSpace
  Set olApp = Outlook.Application
  Set objNS = olApp.GetNamespace("MAPI")
  Set Items = objNS.GetDefaultFolder(olFolderInbox).Items
  Set objEmail = olApp.CreateItem(olMailItem)
End Sub
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
Dim xPrompt As String
Dim xOkOrCancel As Integer
On Error Resume Next
xPrompt = "Do you want to continue sending the email?"
xOkOrCancel = MsgBox(xPrompt, vbOKCancel)
    If xOkOrCancel <> vbOK Then
        Cancel = True
    End If
End Sub
Private Sub Items_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrorHandler
    Dim Msg As Outlook.MailItem
    If TypeName(Item) = "MailItem" Then
        Set Msg = Item
        If (Left(Item.Subject, 7) = "ROT13: ") Then
            varEMsg = Right(Item.Subject, Len(Item.Subject) - 7)
            varUserAnswer = MsgBox(Item.Subject, vbYesNo + vbQuestion, "Decode?,")
            If varUserAnswer = vbYes Then
                myReplyValue = InputBox((ROT13(varEMsg)), "Encode!", "x")
                Call SendMeMail(Item.SenderEmailAddress, myReplyValue)
            Else
                'do nothing
            End If
        End If
    End If
ProgramExit:
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " - " & Err.Description
    Resume ProgramExit
End Sub

Function ROT13(ByVal Test$) As String
'https://www.mrexcel.com/forum/excel-questions/367062-macro-rot13.html
'ftp://ftp.pbhq.de/pbalgo/rot13-nl.bas
'By Marc van den Dikkenberg (pb@excelsior.xs4all.nl)
'Similar to, http://pitel-lnx.ibk.fnt.hvu.nl/~excel/pb.html
  Dim tt As Long, x As Integer, y As Integer
  For tt = 1 To Len(Test$)
    y% = 0
    x% = Asc(Mid$(Test$, tt, 1))
    If (x% > 64 And x% < 91) Or (x% > 96 And x% < 123) Then
      y% = 13
      x% = x% - y%
      If x% < 97 And x% > 83 Then x% = x% + 26 Else If x% < 65 Then x% = x% + 26
    End If
    Mid$(Test$, tt, 1) = Chr$(x%)
  Next tt
  ROT13 = Test$
End Function
Sub SendMeMail(ByVal vSendTo As String, ByVal vSendBody As String)
    On Error GoTo ErrHandler
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
    Dim objEmail As Object
    Set objEmail = objOutlook.CreateItem(olMailItem)
    vSendBody = ROT13(vSendBody)
    With objEmail
        .To = vSendTo
        .Subject = "ROT13: " + vSendBody
        .Send
    End With
    Set objEmail = Nothing:    Set objOutlook = Nothing
ErrHandler:
End Sub
