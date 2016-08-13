Attribute VB_Name = "modText"
Option Explicit

' Text declares
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

' Used to set a font for GDI text drawing
Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
    frmMain.Font = Font
    frmMain.FontSize = Size - 5
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetFont", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' GDI text drawing onto buffer
Public Sub DrawText(ByVal hDC As Long, ByVal x, ByVal y, ByVal text As String, color As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, 0)
    Call TextOut(hDC, x + 1, y + 1, text, Len(text))
    Call SetTextColor(hDC, color)
    Call TextOut(hDC, x, y, text, Len(text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawText", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim Name As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check access level
    If GetPlayerPK(Index) = NO Then

        Select Case GetPlayerAccess(Index)
            Case 0
                color = RGB(255, 96, 0)
            Case 1
                color = QBColor(DarkGrey)
            Case 2
                color = QBColor(Cyan)
            Case 3
                color = QBColor(BrightGreen)
            Case 4
                color = QBColor(Yellow)
        End Select

    Else
        color = QBColor(BrightRed)
    End If

    Name = Trim$(Player(Index).Name)
    ' calc pos
    TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(Name)))
    If GetPlayerSprite(Index) < 1 Or GetPlayerSprite(Index) > NumCharacters Then
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset - 16
    Else
        ' Determine location for text
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset - (DDSD_Character(GetPlayerSprite(Index)).lHeight / 4) + 16
    End If

    ' Draw name
    Call DrawText(TexthDC, TextX, TextY, Name, color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpcName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim Name As String
Dim NpcNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NpcNum = MapNpc(Index).num

    Select Case Npc(NpcNum).Behaviour
        Case NPC_BEHAVIOUR_ATTACKONSIGHT
            color = QBColor(BrightRed)
        Case NPC_BEHAVIOUR_ATTACKWHENATTACKED
            color = QBColor(Yellow)
        Case NPC_BEHAVIOUR_GUARD
            color = QBColor(Grey)
        Case Else
            color = QBColor(BrightGreen)
    End Select

    Name = Trim$(Npc(NpcNum).Name)
    TextX = ConvertMapX(MapNpc(Index).x * PIC_X) + MapNpc(Index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(Name)))
    If Npc(NpcNum).Sprite < 1 Or Npc(NpcNum).Sprite > NumCharacters Then
        TextY = ConvertMapY(MapNpc(Index).y * PIC_Y) + MapNpc(Index).YOffset - 16
    Else
        ' Determine location for text
        TextY = ConvertMapY(MapNpc(Index).y * PIC_Y) + MapNpc(Index).YOffset - (DDSD_Character(Npc(NpcNum).Sprite).lHeight / 4) + 16
    End If

    ' Draw name
    Call DrawText(TexthDC, TextX, TextY, Name, color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpcName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function BltMapAttributes()
    Dim x As Long
    Dim y As Long
    Dim tX As Long
    Dim tY As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optAttribs.Value Then
        For x = TileView.Left To TileView.Right
            For y = TileView.top To TileView.Bottom
                If IsValidMapPoint(x, y) Then
                    With Map.Tile(x, y)
                        tX = ((ConvertMapX(x * PIC_X)) - 4) + (PIC_X * 0.5)
                        tY = ((ConvertMapY(y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                DrawText TexthDC, tX, tY, "B", QBColor(BrightRed)
                            Case TILE_TYPE_WARP
                                DrawText TexthDC, tX, tY, "W", QBColor(BrightBlue)
                            Case TILE_TYPE_ITEM
                                DrawText TexthDC, tX, tY, "I", QBColor(White)
                            Case TILE_TYPE_NPCAVOID
                                DrawText TexthDC, tX, tY, "N", QBColor(White)
                            Case TILE_TYPE_KEY
                                DrawText TexthDC, tX, tY, "K", QBColor(White)
                            Case TILE_TYPE_KEYOPEN
                                DrawText TexthDC, tX, tY, "O", QBColor(White)
                            Case TILE_TYPE_RESOURCE
                                DrawText TexthDC, tX, tY, "O", QBColor(Green)
                            Case TILE_TYPE_DOOR
                                DrawText TexthDC, tX, tY, "D", QBColor(Brown)
                            Case TILE_TYPE_NPCSPAWN
                                DrawText TexthDC, tX, tY, "S", QBColor(Yellow)
                            Case TILE_TYPE_SHOP
                                DrawText TexthDC, tX, tY, "S", QBColor(BrightBlue)
                            Case TILE_TYPE_BANK
                                DrawText TexthDC, tX, tY, "B", QBColor(Blue)
                            Case TILE_TYPE_HEAL
                                DrawText TexthDC, tX, tY, "H", QBColor(BrightGreen)
                            Case TILE_TYPE_TRAP
                                DrawText TexthDC, tX, tY, "T", QBColor(BrightRed)
                            Case TILE_TYPE_SLIDE
                                DrawText TexthDC, tX, tY, "S", QBColor(BrightCyan)
                        End Select
                    End With
                End If
            Next
        Next
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "BltMapAttributes", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub BltActionMsg(ByVal Index As Long)
    Dim x As Long, y As Long, i As Long, Time As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' does it exist
    If ActionMsg(Index).Created = 0 Then Exit Sub

    ' how long we want each message to appear
    Select Case ActionMsg(Index).Type
        Case ACTIONMSG_STATIC
            Time = 1500

            If ActionMsg(Index).y > 0 Then
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) - 2
            Else
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) + 18
            End If

        Case ACTIONMSG_SCROLL
            Time = 1500
        
            If ActionMsg(Index).y > 0 Then
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) - 2 - (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            Else
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) + 18 + (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            End If

        Case ACTIONMSG_SCREEN
            Time = 3000

            ' This will kill any action screen messages that there in the system
            For i = MAX_BYTE To 1 Step -1
                If ActionMsg(i).Type = ACTIONMSG_SCREEN Then
                    If i <> Index Then
                        ClearActionMsg Index
                        Index = i
                    End If
                End If
            Next
            x = (frmMain.picScreen.width \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
            y = 425

    End Select
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)

    If GetTickCount < ActionMsg(Index).Created + Time Then
        Call DrawText(TexthDC, x, y, ActionMsg(Index).message, QBColor(ActionMsg(Index).color))
    Else
        ClearActionMsg Index
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltActionMsg", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function getWidth(ByVal DC As Long, ByVal text As String) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    getWidth = frmMain.TextWidth(text) \ 2
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "getWidth", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub AddText(ByVal Msg As String, ByVal color As Integer)
Dim S As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    S = vbNewLine & Msg
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
    frmMain.txtChat.SelColor = QBColor(color)
    frmMain.txtChat.SelText = S
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AddText", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
