Attribute VB_Name = "modText"
Option Explicit

Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Const Quote As String = """"

Public Const Black As Byte = 0
Public Const blue As Byte = 1
Public Const green As Byte = 2
Public Const Cyan As Byte = 3
Public Const red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15

Public Const SayColor As Byte = White
Public Const GlobalColor As Byte = green
Public Const BroadcastColor As Byte = White
Public Const TellColor As Byte = White
Public Const EmoteColor As Byte = White
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = White
Public Const WhoColor As Byte = Grey
Public Const JoinLeftColor As Byte = Grey
Public Const NpcColor As Byte = White
Public Const AlertColor As Byte = White
Public Const NewMapColor As Byte = Grey

Public TexthDC As Long
Public GameFont As Long

Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
End Sub

Public Sub DrawText(ByVal hDC As Long, ByVal X, ByVal Y, ByVal Text As String, Color As Long)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, RGB(0, 0, 0))
    Call TextOut(hDC, X + 1, Y + 0, Text, Len(Text))
    Call TextOut(hDC, X + 0, Y + 1, Text, Len(Text))
    Call TextOut(hDC, X - 1, Y - 0, Text, Len(Text))
    Call TextOut(hDC, X - 0, Y - 1, Text, Len(Text))
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, X, Y, Text, Len(Text))
End Sub
Public Sub DrawPlayerNameText(ByVal hDC As Long, ByVal X, ByVal Y, ByVal Text As String, Color As Long)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, RGB(0, 0, 0))
    Call TextOut(hDC, X + 1, Y + 0, Text, Len(Text))
    Call TextOut(hDC, X + 0, Y + 1, Text, Len(Text))
    Call TextOut(hDC, X - 1, Y - 0, Text, Len(Text))
    Call TextOut(hDC, X - 0, Y - 1, Text, Len(Text))
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, X, Y, Text, Len(Text))
End Sub

Public Sub DrawTextInter(ByVal hDC As Long, ByVal X, ByVal Y, ByVal Text As String)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, vbBlack)
    Call TextOut(hDC, X, Y, Text, Len(Text))
End Sub

Public Sub AddText(ByVal Msg As String, ByVal Color As Integer)
Dim TmpMsg As String
Dim s As String
Dim C As Long
Dim t As Long
Dim I As Long
Dim J As Integer
Dim z As Long
t = 0
    For J = 0 To ((Len(Msg) \ 50))
        TmpMsg = Mid$(Msg, (50 * J) + 1, 50)

        For I = 1 To MAX_BLT_LINE
             If t = 0 Then
                 If BattlePMsg(I).Index <= 0 Then
                     BattlePMsg(I).Index = 1
                     BattlePMsg(I).Msg = TmpMsg
                     BattlePMsg(I).Color = Color
                     BattlePMsg(I).time = GetTickCount
                     BattlePMsg(I).Done = 1
                     BattlePMsg(I).Y = 0
                     Exit Sub
                 Else
                     BattlePMsg(I).Y = BattlePMsg(I).Y - 15
                 End If
             Else
                 If BattleMMsg(I).Index <= 0 Then
                     BattleMMsg(I).Index = 1
                     BattleMMsg(I).Msg = TmpMsg
                     BattleMMsg(I).Color = Color
                     BattleMMsg(I).time = GetTickCount
                     BattleMMsg(I).Done = 1
                     BattleMMsg(I).Y = 0
                     Exit Sub
                 Else
                     BattleMMsg(I).Y = BattleMMsg(I).Y - 15
                 End If
             End If
        Next I
         
        z = 1
        If t = 0 Then
             For I = 1 To MAX_BLT_LINE
                 If I < MAX_BLT_LINE Then If BattlePMsg(I).Y < BattlePMsg(I + 1).Y Then z = I Else If BattlePMsg(I).Y < BattlePMsg(1).Y Then z = I
             Next I
                         
             BattlePMsg(z).Index = 1
             BattlePMsg(z).Msg = TmpMsg
             BattlePMsg(z).Color = Color
             BattlePMsg(z).time = GetTickCount
             BattlePMsg(z).Done = 1
             BattlePMsg(z).Y = 0
        Else
             For I = 1 To MAX_BLT_LINE
                 If I < MAX_BLT_LINE Then If BattleMMsg(I).Y < BattleMMsg(I + 1).Y Then z = I Else If BattleMMsg(I).Y < BattleMMsg(1).Y Then z = I
             Next I
                         
             BattleMMsg(z).Index = 1
             BattleMMsg(z).Msg = TmpMsg
             BattleMMsg(z).Color = Color
             BattleMMsg(z).time = GetTickCount
             BattleMMsg(z).Done = 1
             BattleMMsg(z).Y = 0
        End If
    Next J
    Exit Sub
End Sub

Public Sub DisplayReport(ByVal Msg As String, ByVal Color As Integer)
    Dim I, z As Integer

    For I = 1 To MAX_BLT_LINE
        If BattleMMsg(I).Index <= 0 Then
            BattleMMsg(I).Index = 1
            BattleMMsg(I).Msg = Msg
            BattleMMsg(I).Color = Color
            BattleMMsg(I).time = GetTickCount
            BattleMMsg(I).Done = 1
            BattleMMsg(I).Y = 0
            Exit Sub
        Else
            BattleMMsg(I).Y = BattleMMsg(I).Y - 15
        End If
    Next I

    z = 1
    For I = 1 To MAX_BLT_LINE
        If I < MAX_BLT_LINE Then If BattleMMsg(I).Y < BattleMMsg(I + 1).Y Then z = I Else If BattleMMsg(I).Y < BattleMMsg(1).Y Then z = I
    Next I
    BattleMMsg(z).Index = 1
    BattleMMsg(z).Msg = Msg
    BattleMMsg(z).Color = Color
    BattleMMsg(z).time = GetTickCount
    BattleMMsg(z).Done = 1
    BattleMMsg(z).Y = 0
End Sub


Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    If NewLine Then Txt.Text = Txt.Text + Msg + vbCrLf Else Txt.Text = Txt.Text + Msg
    
    Txt.SelStart = Len(Txt.Text) - 1
End Sub

Function Parse(ByVal num As Long, ByVal data As String)
Dim I As Long
Dim n As Long
Dim sChar As Long
Dim eChar As Long

    n = 0
    sChar = 1
    
    For I = 1 To Len(data)
        If Mid$(data, I, 1) = SEP_CHAR Then
            If n = num Then
                eChar = I
                Parse = Mid$(data, sChar, eChar - sChar)
                Exit For
            End If
            
            sChar = I + 1
            n = n + 1
        End If
    Next I
    
End Function

