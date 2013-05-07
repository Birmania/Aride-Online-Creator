Attribute VB_Name = "modText"
Option Explicit
'faire mousedown pour selectioner plusieurs careaus de tiles
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal o As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
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

Public Const SayColor As Byte = Grey
Public Const GlobalColor As Byte = Green
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

Public Sub DrawText(ByVal hDC As Long, ByVal x, ByVal y, ByVal Text As String, Color As Long)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, RGB(0, 0, 0))
    Call TextOut(hDC, x + 0, y + 0, Text, Len(Text))
    Call TextOut(hDC, x + 1, y + 1, Text, Len(Text))
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, x, y, Text, Len(Text))
End Sub
Public Sub DrawPlayerNameText(ByVal hDC As Long, ByVal x, ByVal y, ByVal Text As String, Color As Long)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, RGB(0, 0, 0))
    Call TextOut(hDC, x + 0, y + 0, Text, Len(Text))
    Call TextOut(hDC, x + 1, y + 1, Text, Len(Text))
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, x, y, Text, Len(Text))
End Sub

Public Sub AddText(ByVal Msg As String, ByVal Color As Long)
Dim s As String
Dim c As Long
Dim t As Long
Dim I As Long
Dim z As Long
t = 0
       For I = 1 To MAX_BLT_LINE
            If t = 0 Then
                If BattlePMsg(I).Index <= 0 Then
                    BattlePMsg(I).Index = 1
                    BattlePMsg(I).Msg = Msg
                    BattlePMsg(I).Color = Color
                    BattlePMsg(I).Time = GetTickCount
                    BattlePMsg(I).Done = 1
                    BattlePMsg(I).y = 0
                    Exit Sub
                Else
                    BattlePMsg(I).y = BattlePMsg(I).y - 15
                End If
            Else
                If BattleMMsg(I).Index <= 0 Then
                    BattleMMsg(I).Index = 1
                    BattleMMsg(I).Msg = Msg
                    BattleMMsg(I).Color = Color
                    BattleMMsg(I).Time = GetTickCount
                    BattleMMsg(I).Done = 1
                    BattleMMsg(I).y = 0
                    Exit Sub
                Else
                    BattleMMsg(I).y = BattleMMsg(I).y - 15
                End If
            End If
        Next I
        
        z = 1
        If t = 0 Then
            For I = 1 To MAX_BLT_LINE
                If I < MAX_BLT_LINE Then If BattlePMsg(I).y < BattlePMsg(I + 1).y Then z = I Else If BattlePMsg(I).y < BattlePMsg(1).y Then z = I
            Next I
                        
            BattlePMsg(z).Index = 1
            BattlePMsg(z).Msg = Msg
            BattlePMsg(z).Color = Color
            BattlePMsg(z).Time = GetTickCount
            BattlePMsg(z).Done = 1
            BattlePMsg(z).y = 0
        Else
            For I = 1 To MAX_BLT_LINE
                If I < MAX_BLT_LINE Then If BattleMMsg(I).y < BattleMMsg(I + 1).y Then z = I Else If BattleMMsg(I).y < BattleMMsg(1).y Then z = I
            Next I
                        
            BattleMMsg(z).Index = 1
            BattleMMsg(z).Msg = Msg
            BattleMMsg(z).Color = Color
            BattleMMsg(z).Time = GetTickCount
            BattleMMsg(z).Done = 1
            BattleMMsg(z).y = 0
        End If
        Exit Sub
End Sub

Function Parse(ByVal num As Long, ByVal Data As String)
Dim I As Long
Dim n As Long
Dim sChar As Long
Dim eChar As Long

    n = 0
    sChar = 1
    
    For I = 1 To Len(Data)
        If Mid$(Data, I, 1) = SEP_CHAR Then
            If n = num Then eChar = I: Parse = Mid$(Data, sChar, eChar - sChar): Exit For
            
            sChar = I + 1
            n = n + 1
        End If
    Next I
End Function

