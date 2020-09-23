VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.UserControl ActionTree 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   MouseIcon       =   "ActionTree.ctx":0000
   PaletteMode     =   4  'Nichts
   ScaleHeight     =   394
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   410
   Begin MSComCtl2.FlatScrollBar FlatScrollBar3 
      Height          =   5655
      Left            =   5880
      TabIndex        =   3
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   9975
      _Version        =   393216
      Orientation     =   1245184
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar2 
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   5640
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1245185
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5640
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1245185
   End
   Begin VB.PictureBox Cutter 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4320
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "ActionTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum eSlotState
    
    essClose = 0
    essopen = 1
    essSelected = 2
    
End Enum

Private Type SlotItem

    iSlotTop As Single
    iSlotContent As String
    iLevel As Integer

End Type

Private Type ColumnSize
    
    iColIndex As Integer
    iColLeft As Integer

End Type

Private Type SlotItemNode

    iSlotContent As String
    iSlotState As eSlotState
    iSlotChildCount As Integer

End Type

Private Type iRECT

    Top As Integer
    Left As Integer
    Width As Integer
    Height As Integer

End Type

Private Const ICON_LEFT_OFFSET As Integer = 16

Public AT_Nodes As Nodes
Public AT_Columns As Columns

Dim iSlot() As SlotItem
Dim iSlotNode() As SlotItemNode
Dim iFirstItem As Integer
Dim iItemHeight As Integer
Dim iColHeight As Integer
Dim iTreeWidth As Integer
Dim iLeftOffset As Integer
Dim iSelColBar As ColumnSize

'Ereignisdeklarationen:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Tritt auf, wenn der Benutzer eine Maustaste über einem Objekt drückt und wieder losläßt."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Tritt auf, wenn der Benutzer eine Maustaste über einem Objekt drückt und wieder losläßt und anschließend erneut drückt und wieder losläßt."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Tritt auf, wenn der Benutzer eine Taste drückt, während ein Objekt den Fokus besitzt."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Tritt auf, wenn der Benutzer eine ANSI-Taste drückt und losläßt."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Tritt auf, wenn der Benutzer eine Taste losläßt, während ein Objekt den Fokus hat."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste drückt, während ein Objekt den Fokus hat."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Tritt auf, wenn der Benutzer die Maus bewegt."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste losläßt, während ein Objekt den Fokus hat."

Event ItemClick(Button As Integer, Shift As Integer, NodeItem As Node)

'Standard-Eigenschaftswerte:
Const m_def_SelectedTextColor = -2147483643
Const m_def_SelectedItem = 0
Const m_def_LightColor = 12632256
Const m_def_Darkcolor = 8421504
Const m_def_Selectioncolor = 8388608
'Eigenschaftsvariablen:
Dim m_Icons As Object
Dim m_SelectedTextColor As Long
Dim m_SelectedItem As Integer
Dim m_LightColor As Variant
Dim m_Darkcolor As Long
Dim m_Selectioncolor As Variant

Private Function TEXT_LEFT_OFFSET() As Integer
    
    If m_Icons Is Nothing Then
        TEXT_LEFT_OFFSET = 16
    Else
        TEXT_LEFT_OFFSET = ICON_LEFT_OFFSET + m_Icons.ImageWidth + 5
    End If

End Function

Private Function SEL_LEFT_OFFSET() As Integer
    SEL_LEFT_OFFSET = TEXT_LEFT_OFFSET - 2
End Function

Public Sub Clear_Items()
    
    Dim i As Integer
    
    If AT_Nodes.Count = 0 Then Exit Sub
    
    For i = 1 To AT_Nodes.Count
        AT_Nodes.Remove 1
    Next i
    
    Refresh
    
End Sub

Public Sub Clear_Columns()
    
    Dim i As Integer
    
    If AT_Columns.Count = 0 Then Exit Sub
    
    For i = 1 To AT_Columns.Count
        AT_Columns.Remove 1
    Next i
    
    Refresh
    
End Sub

Public Sub Refresh()
    
    Dim i As Integer
    Dim c As Integer
    ReDim iSlotNode(AT_Nodes.Count)
    
    iLeftOffset = 0
    Define_Slots
    For i = 1 To AT_Nodes.Count
    
        iSlotNode(i).iSlotContent = AT_Nodes.Item(i).Key
        iSlotNode(i).iSlotChildCount = 0
        iSlotNode(i).iSlotState = essClose
        
    Next i
    
    'set the child count
    For i = 1 To AT_Nodes.Count
        
        If AT_Nodes(i).Releative <> "" Then
                
            c = iSlotNode(GetParentIndex(AT_Nodes(i).Releative)).iSlotChildCount
            c = c + 1
            iSlotNode(GetParentIndex(AT_Nodes(i).Releative)).iSlotChildCount = c
        
        End If
    
    Next i
    
    
    Draw
    
End Sub

Private Function GetParentIndex(sParentKey As String) As Integer

    Dim i As Integer
    
    
    For i = 1 To AT_Nodes.Count
        If AT_Nodes.Item(i).Key = sParentKey Then
            GetParentIndex = i
            Exit Function
        End If
    Next i

End Function

Private Function iItemCount() As Integer
    Dim ic As Integer
    ic = UBound(iSlot)
    If ic > AT_Nodes.Count Then ic = AT_Nodes.Count
    iItemCount = ic
End Function

Private Sub Draw_Col()
    
    If iColHeight = 0 Then Exit Sub
    If AT_Columns.Count = 0 Then Exit Sub
    
    Dim i As Integer
    Dim iL As Integer
    
    Cutter.Width = UserControl.ScaleWidth
    Cutter.Height = iColHeight
    
    Cutter.BackColor = LightColor
    Cutter.Line (0, 0)-(Cutter.Width - 1, 0), Darkcolor
    Cutter.Line (0, 0)-(0, Cutter.Height), Darkcolor
    Cutter.Line (0, Cutter.Height - 1)-(Cutter.Width - 1, Cutter.Height - 1), Darkcolor
    Cutter.Line (Cutter.Width - 1, 0)-(Cutter.Width - 1, Cutter.Height), Darkcolor
    
    UserControl.PaintPicture Cutter.Image, 0, 0
    
    iL = 5
    
    UserControl.FontBold = True
    
    For i = 1 To AT_Columns.Count
    
        UserControl.CurrentY = ((iColHeight / 2) - (UserControl.TextHeight("I") / 2))
        UserControl.CurrentX = iL
        UserControl.Print ShortWord(AT_Columns(i).Caption, AT_Columns(i).ColWidth - 5)
        
        iL = iL + AT_Columns(i).ColWidth
        
        UserControl.Line (iL - 5, 0)-(iL - 5, iColHeight), Darkcolor
        UserControl.Line (iL - 5, iColHeight)-(iL - 5, UserControl.Height), LightColor
        
    Next i
    
    UserControl.FontBold = False

End Sub

Private Sub Draw(Optional iScrollOnly As Boolean = False)
    
    UserControl.Cls
    
    
    If AT_Nodes Is Nothing Then Exit Sub
    If AT_Columns Is Nothing Then Exit Sub
    If AT_Nodes.Count = 0 Then
        iFirstItem = 1
        iTreeWidth = 0
        'AsignScrollbars
        'DrawScrollBarPatch
        FlatScrollBar1.Visible = False
        FlatScrollBar2.Visible = False
        FlatScrollBar3.Visible = False
        Exit Sub
    End If
    
    Dim i As Integer
    Dim i2 As Integer
    Dim ic As Integer
    Dim ctr As Integer
    Dim oL As Long
    
    Draw_Col
    
    ic = iItemCount
    ctr = 1
    oL = UserControl.ForeColor
    iTreeWidth = 0
    
    For i = iFirstItem To AT_Nodes.Count 'ic
        If AT_Nodes(i).Releative = "" Then
        
            UserControl.ForeColor = oL
            
            If SelectedItem = i Then
                DrawSelection iSlot(ctr).iSlotTop * 1, GetLevelLeft(1) + SEL_LEFT_OFFSET, Selectioncolor, ctr
                UserControl.ForeColor = SelectedTextColor
            End If
            
            
            iSlot(ctr).iSlotContent = AT_Nodes.Item(i).Key
            iSlot(ctr).iLevel = 1
            
'            UserControl.CurrentY = 3
'            UserControl.CurrentX = TEXT_LEFT_OFFSET
'            UserControl.Print
            
            SetHigherWidth GetLevelLeft(1) + UserControl.TextWidth(AT_Nodes.Item(i).Caption)
            
            WriteText iSlot(ctr).iSlotTop * 1, GetLevelLeft(1), ShortWord(AT_Nodes.Item(i).Caption, Get_ColWidth(1) - GetLevelLeft(1) - TEXT_LEFT_OFFSET)
            
            Draw_Line ctr
            DrawIcon iSlot(ctr).iSlotTop * 1, GetLevelLeft(1), AT_Nodes.Item(i).Icon
            
            DrawArrows ctr, i, 1
            UserControl.ForeColor = oL
            DrawNodeChilds ctr, i
            
            ctr = ctr + 1
            
            If ctr > ic Then
                If iScrollOnly = False Then AsignScrollbars
                DrawScrollBarPatch
                Exit Sub
            End If
            
            
            If (iSlotNode(i).iSlotChildCount <> 0) And (iSlotNode(i).iSlotState = essopen) Then AddSlotItem ctr, AT_Nodes.Item(i).Key, ic, , AT_Nodes.Item(i).Family
            If ctr > ic Then
                If iScrollOnly = False Then AsignScrollbars
                DrawScrollBarPatch
                Exit Sub
            End If
        
        End If
        
    Next i

    If iScrollOnly = False Then AsignScrollbars
    DrawScrollBarPatch

End Sub

Public Function HasChildren() As Boolean

    If iSlotNode(SelectedItem).iSlotChildCount <> 0 Then
        HasChildren = True
    Else
        HasChildren = False
    End If

End Function

Private Sub DrawScrollBarPatch()
        
    If (FlatScrollBar2.Visible = True And FlatScrollBar3.Visible = True) Or ((FlatScrollBar1.Visible = True And FlatScrollBar3.Visible = True) And AT_Columns.Count = 0) Then
    
        Cutter.Cls
        Cutter.Width = FlatScrollBar3.Width
        Cutter.Height = FlatScrollBar1.Height
        Cutter.BackColor = LightColor
        UserControl.PaintPicture Cutter.Image, FlatScrollBar3.Left, FlatScrollBar3.Top + FlatScrollBar3.Height
    
    End If

End Sub

Private Sub AddSlotItem(SlotCounter As Integer, sParentKey As String, iSlotMax As Integer, Optional iLevel As Integer = 2, Optional sFamily As String = "")

    Dim i As Integer
    Dim oL As Long
    
    oL = UserControl.ForeColor

    For i = 1 To AT_Nodes.Count
        
        If AT_Nodes(i).Releative = sParentKey Then
        
            UserControl.ForeColor = oL
            If SelectedItem = i Then
                DrawSelection iSlot(SlotCounter).iSlotTop * 1, GetLevelLeft(iLevel) + SEL_LEFT_OFFSET, Selectioncolor, SlotCounter
                UserControl.ForeColor = SelectedTextColor
            End If
            
            
            If SlotCounter > iSlotMax Then: UserControl.ForeColor = oL: Exit Sub
            iSlot(SlotCounter).iSlotContent = AT_Nodes.Item(i).Key
            iSlot(SlotCounter).iLevel = iLevel
            
'            UserControl.CurrentY = iSlot(SlotCounter).iSlotTop + 3
'            UserControl.CurrentX = GetLevelLeft(iLevel) + TEXT_LEFT_OFFSET
'            UserControl.Print AT_Nodes.Item(i).Caption
            
            SetHigherWidth GetLevelLeft(iLevel) + UserControl.TextWidth(AT_Nodes.Item(i).Caption)
            WriteText iSlot(SlotCounter).iSlotTop * 1, GetLevelLeft(iLevel), ShortWord(AT_Nodes.Item(i).Caption, Get_ColWidth(1) - GetLevelLeft(iLevel) - TEXT_LEFT_OFFSET)
            
            Draw_Line SlotCounter
            DrawIcon iSlot(SlotCounter).iSlotTop * 1, GetLevelLeft(iLevel), AT_Nodes.Item(i).Icon
            DrawArrows SlotCounter, i, iLevel
            UserControl.ForeColor = oL
            DrawNodeChilds SlotCounter, i
            SlotCounter = SlotCounter + 1
            
            If SlotCounter > iSlotMax Then: UserControl.ForeColor = oL: Exit Sub
            
            If (iSlotNode(i).iSlotChildCount <> 0) And (iSlotNode(i).iSlotState = essopen) Then
                iLevel = iLevel + 1
                AddSlotItem SlotCounter, AT_Nodes(i).Key, iSlotMax, iLevel, sFamily
                iLevel = iLevel - 1
            End If
            
        End If
        
    Next i
    

End Sub

Private Sub SetHigherWidth(iNewWidth As Integer)

    Dim i As Integer
    
    If FlatScrollBar3.Visible = True And AT_Columns.Count = 0 Then
        i = iNewWidth + TEXT_LEFT_OFFSET + 5 + FlatScrollBar3.Width
    Else
        i = iNewWidth + TEXT_LEFT_OFFSET + 5
    End If

        
    If i > iTreeWidth Then iTreeWidth = i

End Sub

Private Function Get_ColWidth(iColIndex As Integer) As Integer

    If AT_Columns.Count = 0 Then
        If FlatScrollBar3.Visible = True Then
            Get_ColWidth = UserControl.Width - FlatScrollBar3.Width
        Else
            Get_ColWidth = UserControl.Width
        End If
    Else
        Get_ColWidth = AT_Columns(iColIndex).ColWidth
    End If

End Function

Private Function ShortWord(sString As String, RefLenght As Long) As String
    
    Dim tmpString As String
    Dim i As Integer
    
    
    i = InStr(1, sString, vbCr)
    If i <> 0 Then
        tmpString = Mid(sString, 1, i)
    Else
        tmpString = sString
    End If
    
    If UserControl.TextWidth(tmpString) > RefLenght Then
    
        Do Until UserControl.TextWidth(tmpString) < RefLenght
        
            If Len(tmpString) <= 3 Then: tmpString = "": Exit Do
            tmpString = Left(tmpString, Len(tmpString) - 4) & "..."
            
        Loop
        
    End If
    
    ShortWord = tmpString

End Function

Private Function GetLevelLeft(iIndex As Integer) As Integer

    '3
    '13
    '32
    
    If iIndex = 1 Then
        GetLevelLeft = 3 - iLeftOffset
    Else
        GetLevelLeft = (32 * (iIndex - 1)) - iLeftOffset
    End If

End Function

Private Sub DrawArrows(SlotCounter As Integer, iNodeIndex As Integer, iLevel As Integer)

    Dim st As String
    Dim iSize As Integer
    Dim lForeColor As Long
    Dim iLl As Integer
    
    st = UserControl.FontName
    iSize = UserControl.FontSize
    lForeColor = UserControl.ForeColor
    
    If iSlotNode(iNodeIndex).iSlotChildCount <> 0 Then
    
        iLl = GetLevelLeft(iLevel)
        
        If AT_Columns.Count <> 0 Then
            If iLl > AT_Columns(1).ColWidth - 9 Then Exit Sub
        End If
    
        UserControl.CurrentY = iSlot(SlotCounter).iSlotTop + ((iItemHeight / 2) - 7)
        UserControl.CurrentX = iLl
        UserControl.FontName = "webdings"
        UserControl.FontSize = 7
        UserControl.ForeColor = Darkcolor
        UserControl.Print "c"
        
        UserControl.CurrentY = iSlot(SlotCounter).iSlotTop + ((iItemHeight / 2) - 5)
        UserControl.FontName = "tahoma"
        UserControl.FontSize = 7.5
        UserControl.ForeColor = vbBlack
        
        Select Case iSlotNode(iNodeIndex).iSlotState
            Case essClose
                UserControl.CurrentX = GetLevelLeft(iLevel) + 1
                UserControl.Print "+"
            Case essopen
                UserControl.CurrentX = GetLevelLeft(iLevel) + 2
                UserControl.Print "–"
            
        End Select
        
        UserControl.FontName = st
        UserControl.FontSize = iSize
        UserControl.ForeColor = lForeColor
        
    End If
    
End Sub

Private Sub WriteText(Y As Integer, X As Integer, sString As String)

    UserControl.CurrentY = Y + ((iItemHeight / 2) - (UserControl.TextHeight("I") / 2))
    UserControl.CurrentX = X + TEXT_LEFT_OFFSET
    UserControl.Print sString

End Sub

Private Sub DrawIcon(Y As Integer, X As Integer, IconNum As Integer)
    If IconNum = 0 Then Exit Sub
    If m_Icons Is Nothing Then Exit Sub
    If IconNum > m_Icons.ListImages.Count Then Exit Sub
    
    Dim i As Integer
    
    i = Get_ColWidth(1) - (X + ICON_LEFT_OFFSET)
    If i <= 0 Then Exit Sub
    UserControl.PaintPicture m_Icons.ListImages(IconNum).ExtractIcon, X + ICON_LEFT_OFFSET, Y + ((iItemHeight / 2) - (m_Icons.ImageHeight / 2)), , , , , i
End Sub

Private Sub DrawSelection(Y As Integer, X As Integer, lBackColor As Long, iSlotNum As Integer)

    Dim lLineColor As Long
    Dim iW As Integer

    Cutter.Height = iItemHeight
    If AT_Columns.Count = 0 Then
        iW = UserControl.ScaleWidth - X
        If iW <= 1 Then Exit Sub
        Cutter.Width = iW
        Cutter.BackColor = lBackColor
        UserControl.PaintPicture Cutter.Image, X, Y
    Else
        iW = AT_Columns.Item(1).ColWidth - X
        If iW > 1 Then
            Cutter.Width = iW
            Cutter.BackColor = lBackColor
            UserControl.PaintPicture Cutter.Image, X, Y
        End If
        
        iW = UserControl.ScaleWidth - AT_Columns.Item(1).ColWidth
        If iW <= 1 Then Exit Sub
        Cutter.Width = iW
        Cutter.BackColor = LightColor
        UserControl.PaintPicture Cutter.Image, AT_Columns.Item(1).ColWidth, Y
        
    End If
    

End Sub

Private Sub Draw_Line(i_SlotIndex As Integer)
    Dim lForeColor As Long
    
    lForeColor = UserControl.ForeColor
    UserControl.ForeColor = LightColor
    
    UserControl.Line (0, iSlot(i_SlotIndex).iSlotTop)-(UserControl.ScaleWidth, iSlot(i_SlotIndex).iSlotTop)
    UserControl.Line (0, iSlot(i_SlotIndex).iSlotTop + iItemHeight)-(UserControl.ScaleWidth, iSlot(i_SlotIndex).iSlotTop + iItemHeight)
    
    UserControl.ForeColor = lForeColor

End Sub

Private Sub Define_Slots()

    Dim i As Single
    Dim tH As Integer
    Dim iH As Integer
    Dim sH As Integer
    
    If m_Icons Is Nothing Then
        tH = 16
    Else
        tH = m_Icons.ImageHeight
    End If
    
    iH = UserControl.TextHeight("I")
    
    If tH > iH Then
        iItemHeight = tH + 6
    Else
        iItemHeight = iH + 6
    End If
     
    If AT_Columns.Count <> 0 Then
        iColHeight = iH + 6
    Else
        iColHeight = 0
    End If
    
    sH = UserControl.ScaleHeight
    If sH < iItemHeight Then sH = iItemHeight
    i = Format(sH / iItemHeight, "##")
    
    If (iItemHeight * i < UserControl.ScaleHeight) Then i = i + 1
    'If (iItemHeight * i > UserControl.ScaleHeight) Then i = i - 1
    
    ReDim iSlot(i)
    
    For i = 1 To UBound(iSlot)
        iSlot(i).iSlotTop = ((iItemHeight * i) - iItemHeight) + iColHeight
    Next i

End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Draw
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gibt die Vordergrundfarbe zurück, die zum Anzeigen von Text und Grafiken in einem Objekt verwendet wird, oder legt diese fest."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    Draw
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gibt einen Wert zurück, der bestimmt, ob ein Objekt auf vom Benutzer erzeugte Ereignisse reagieren kann, oder legt diesen fest."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    Draw
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Gibt ein Font-Objekt zurück."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Define_Slots
    Draw
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Gibt den Rahmenstil für ein Objekt zurück oder legt diesen fest."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
    Define_Slots
    Draw
End Property

Private Sub FlatScrollBar1_Change()
    
    Dim i As Integer
    
    i = FlatScrollBar1.Value
    If i < 0 Then i = i * -1
    iLeftOffset = i
    
    Draw True
    
End Sub

Private Sub FlatScrollBar1_Scroll()
    FlatScrollBar1_Change
End Sub

Private Sub FlatScrollBar3_Change()

    iFirstItem = FlatScrollBar3.Value
    Draw True
    
End Sub

Private Sub FlatScrollBar3_Scroll()
    FlatScrollBar3_Change
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()

    Set AT_Nodes = New Nodes
    Set AT_Columns = New Columns
    iFirstItem = 1

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If iSelColBar.iColIndex <> 0 Then
        iSelColBar.iColIndex = 0
        MousePointer = 0
        Draw
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    Dim i As Integer
    Dim ic As Integer
    
    If iSelColBar.iColIndex <> 0 Then
        
        i = X - iSelColBar.iColLeft
        If i <= 16 Then i = 16
        
        AT_Columns(iSelColBar.iColIndex).ColWidth = i
        Draw True
        
    End If
    
    If (Y < iColHeight) Then
        If AT_Columns.Count <> 0 Then
            For i = 1 To AT_Columns.Count
                ic = ic + AT_Columns(i).ColWidth
                If (X > ic - 5) And (X < ic + 5) Then
                    UserControl.MousePointer = 99
                    Exit Sub
                End If
            Next i
        End If
    End If
    
    UserControl.MousePointer = 0
    
End Sub

Public Sub AutoSetColWidth(iColIndex As Integer)

    Dim i As Integer
    Dim i2 As Integer
    Dim iL As Integer
    Dim cL As Integer
    
    On Error Resume Next
    
    iL = 0
    If iColIndex = 0 Then Exit Sub
    
    If iColIndex = 1 Then
        iL = iTreeWidth
    Else
        For i = 1 To UBound(iSlot)
        
            i2 = GetNodePropIndex(i)
            
            If i2 <> 0 Then
            
                If AT_Nodes(i2).Child.Count >= (iColIndex - 1) Then
                    cL = UserControl.TextWidth(AT_Nodes(i2).Child(iColIndex - 1).Caption)
                    If iL < cL Then iL = cL
                End If
            
            End If
                
        Next i
    End If
    
    AT_Columns(iColIndex).ColWidth = iL + 16
    Draw

End Sub

Private Sub UserControl_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
    Dim i As Integer
    Dim ic As Integer
    Dim iCross As Boolean
    
    ic = 0
    If (Y < iColHeight) Then
        If AT_Columns.Count <> 0 Then
            For i = 1 To AT_Columns.Count
                ic = ic + AT_Columns(i).ColWidth
                If (X > ic - 5) And (X < ic + 5) Then
                    
                    If Button = 2 Then
                        AutoSetColWidth i
                        Exit Sub
                    End If
                    
                    iSelColBar.iColIndex = i
                    iSelColBar.iColLeft = ic - AT_Columns(i).ColWidth
                    
                    Exit For
                End If
            Next i
        End If
        Draw True
        Exit Sub
    End If
    
    ic = iItemCount
    
    For i = 1 To ic
            
            If AT_Columns.Count > 0 Then
                If X > AT_Columns(1).ColWidth Then
                    iCross = False
                Else
                    iCross = True
                End If
            Else
                iCross = True
            End If
                
            If iCross = True Then
                If (X > GetLevelLeft(iSlot(i).iLevel)) And (X < GetLevelLeft(iSlot(i).iLevel) + UserControl.TextWidth("W")) Then
                
                    If (Y > iSlot(i).iSlotTop) And (Y < iSlot(i).iSlotTop + iItemHeight) And (iSlotNode(GetNodePropIndex(i)).iSlotChildCount <> 0) Then
                        If iSlotNode(GetNodePropIndex(i)).iSlotState = essopen Then
                            iSlotNode(GetNodePropIndex(i)).iSlotState = essClose
                        ElseIf iSlotNode(GetNodePropIndex(i)).iSlotState = essClose Then
                            iSlotNode(GetNodePropIndex(i)).iSlotState = essopen
                        End If
                        Draw
                        Exit Sub
                    End If
                    
                End If
            End If
            
            If (Y > iSlot(i).iSlotTop) And (Y < iSlot(i).iSlotTop + iItemHeight) And (X > GetLevelLeft(iSlot(i).iLevel) + UserControl.TextWidth("W")) Then
                SelectedItem = GetNodePropIndex(i)
                Draw True
                If SelectedItem <> 0 Then _
                    RaiseEvent ItemClick(Button, Shift, AT_Nodes.Item(SelectedItem))
            End If
            
    Next i
    
    'Draw True
    
End Sub

Private Sub DrawNodeChilds(iSlotNumber As Integer, iItemIndex As Integer)

    Dim i As Integer
    Dim i2 As Integer
    Dim iL As Integer
    
    If AT_Columns.Count <= 1 Then Exit Sub
    
    i2 = AT_Columns.Count
    If AT_Nodes(iItemIndex).Child.Count < i2 Then i2 = AT_Nodes(iItemIndex).Child.Count
    
    iL = AT_Columns(1).ColWidth + 4
    
    For i = 1 To i2
        
        UserControl.CurrentY = iSlot(iSlotNumber).iSlotTop + ((iItemHeight / 2) - (UserControl.TextHeight("I") / 2))
        UserControl.CurrentX = iL
        UserControl.Print ShortWord(AT_Nodes.Item(iItemIndex).Child(i).Caption, AT_Columns(i + 1).ColWidth - 4)

        iL = iL + AT_Columns(i + 1).ColWidth
    Next i
    
End Sub

Private Function GetNodePropIndex(iSlotNumber As Integer) As Integer

    Dim i As Integer
    
    For i = 1 To AT_Nodes.Count
    
        If iSlot(iSlotNumber).iSlotContent = AT_Nodes(i).Key Then
            GetNodePropIndex = i
            Exit Function
        End If
    
    Next i

End Function

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_LightColor = m_def_LightColor
    m_Darkcolor = m_def_Darkcolor
    m_Selectioncolor = m_def_Selectioncolor
    m_SelectedItem = m_def_SelectedItem
    m_SelectedTextColor = m_def_SelectedTextColor
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_LightColor = PropBag.ReadProperty("LightColor", m_def_LightColor)
    m_Darkcolor = PropBag.ReadProperty("Darkcolor", m_def_Darkcolor)
    m_Selectioncolor = PropBag.ReadProperty("Selectioncolor", m_def_Selectioncolor)
    m_SelectedItem = PropBag.ReadProperty("SelectedItem", m_def_SelectedItem)
    m_SelectedTextColor = PropBag.ReadProperty("SelectedTextColor", m_def_SelectedTextColor)
    'Set m_Icons = PropBag.ReadProperty("Icons", Nothing)
End Sub

Private Sub AsignScrollbars()

    Dim Sb1h As Boolean
    Dim Sb2h As Boolean
    Dim Sb3h As Boolean
    
    Dim i As Integer
    Dim ic As Integer
    
    On Error Resume Next
    
    FlatScrollBar1.Visible = False
    FlatScrollBar2.Visible = False
    FlatScrollBar3.Visible = False
    
    If AT_Columns.Count = 0 Then
    
        If iTreeWidth + iLeftOffset > UserControl.ScaleWidth Then
            Sb1h = True
        Else
            Sb1h = False
        End If
        
        Sb2h = False
        
    Else
    
        If iTreeWidth + iLeftOffset > AT_Columns(1).ColWidth Then
            Sb1h = True
        Else
            Sb1h = False
        End If
        
        If AT_Columns.Count > 1 Then
            
            ic = 2
            For i = 1 To AT_Columns.Count
                ic = AT_Columns(i).ColWidth + ic
            Next i
            
            If ic > UserControl.ScaleWidth Then
                Sb2h = True
            Else
                Sb2h = False
            End If
        
        Else
            Sb2h = False
        End If
        
        If AT_Columns(1).ColWidth < 10 Then Sb1h = False
        If UserControl.ScaleWidth - AT_Columns(1).ColWidth < 10 Then Sb2h = False
    
    End If
    
    If (iItemCount * iItemHeight) > UserControl.ScaleHeight Then
        Sb3h = True
    Else
        Sb3h = False
    End If
    
    If UserControl.ScaleHeight <= iColHeight + FlatScrollBar1.Height Or UserControl.ScaleHeight <= iColHeight + FlatScrollBar2.Height Then: Sb1h = False: Sb2h = False
    If UserControl.ScaleHeight - iColHeight < 10 Then Sb3h = False
    
    If AT_Columns.Count = 0 Then
    
        If Sb1h = True Then
            FlatScrollBar1.Left = 0
            If Sb3h = True Then
                FlatScrollBar1.Width = UserControl.ScaleWidth - FlatScrollBar3.Width
            Else
                FlatScrollBar1.Width = UserControl.ScaleWidth
            End If
            FlatScrollBar1.Top = UserControl.ScaleHeight - FlatScrollBar1.Height
            FlatScrollBar1.Min = 0
            FlatScrollBar1.Max = iTreeWidth - UserControl.ScaleWidth
            FlatScrollBar1.Visible = True
        Else
            FlatScrollBar1.Visible = False
            iLeftOffset = 0
            Draw True
        End If
        
        FlatScrollBar2.Visible = False
        
    Else
    
        If Sb1h = True Then
            FlatScrollBar1.Left = 0
            FlatScrollBar1.Width = AT_Columns(1).ColWidth
            FlatScrollBar1.Top = UserControl.ScaleHeight - FlatScrollBar1.Height
            FlatScrollBar1.Min = 0
            FlatScrollBar1.Max = iTreeWidth - AT_Columns(1).ColWidth
            FlatScrollBar1.Visible = True
        Else
            FlatScrollBar1.Visible = False
            iLeftOffset = 0
            Draw True
        End If
    
        If Sb2h = True Then
            FlatScrollBar2.Left = AT_Columns(1).ColWidth + 1
            FlatScrollBar2.Top = UserControl.ScaleHeight - FlatScrollBar2.Height
            If Sb3h = True Then
                FlatScrollBar2.Width = (UserControl.ScaleWidth - AT_Columns(1).ColWidth) - FlatScrollBar3.Width
            Else
                FlatScrollBar2.Width = UserControl.ScaleWidth - AT_Columns(1).ColWidth - 1
            End If
            FlatScrollBar2.Visible = True
        Else
            FlatScrollBar2.Visible = False
        End If
        
    End If
    
    If Sb3h = True Then
        FlatScrollBar3.Top = iColHeight
        FlatScrollBar3.Left = UserControl.ScaleWidth - FlatScrollBar3.Width
        If Sb2h = True Or (AT_Columns.Count = 0 And Sb1h = True) Then
            FlatScrollBar3.Height = UserControl.ScaleHeight - FlatScrollBar3.Top - FlatScrollBar2.Height
        Else
            FlatScrollBar3.Height = UserControl.ScaleHeight - FlatScrollBar3.Top
        End If
        FlatScrollBar3.Min = 1
        FlatScrollBar3.Max = AT_Nodes.Count
        FlatScrollBar3.Visible = True
    Else
        FlatScrollBar3.Visible = False
    End If

End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    Define_Slots
    Draw
    
End Sub

Private Sub UserControl_Terminate()

     If Not (AT_Nodes Is Nothing) Then Set AT_Nodes = Nothing
     If Not (AT_Columns Is Nothing) Then Set AT_Columns = Nothing

End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("LightColor", m_LightColor, m_def_LightColor)
    Call PropBag.WriteProperty("Darkcolor", m_Darkcolor, m_def_Darkcolor)
    Call PropBag.WriteProperty("Selectioncolor", m_Selectioncolor, m_def_Selectioncolor)
    Call PropBag.WriteProperty("SelectedItem", m_SelectedItem, m_def_SelectedItem)
    Call PropBag.WriteProperty("SelectedTextColor", m_SelectedTextColor, m_def_SelectedTextColor)
    
    'Call PropBag.WriteProperty("Icons", m_Icons, Nothing)
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=14,0,0,12632256
Public Property Get LightColor() As Variant
    LightColor = m_LightColor
End Property

Public Property Let LightColor(ByVal New_LightColor As Variant)
    m_LightColor = New_LightColor
    PropertyChanged "LightColor"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=8,0,0,0
Public Property Get Darkcolor() As Long
    Darkcolor = m_Darkcolor
End Property

Public Property Let Darkcolor(ByVal New_Darkcolor As Long)
    m_Darkcolor = New_Darkcolor
    PropertyChanged "Darkcolor"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=14,0,0,8388608
Public Property Get Selectioncolor() As Variant
    Selectioncolor = m_Selectioncolor
End Property

Public Property Let Selectioncolor(ByVal New_Selectioncolor As Variant)
    m_Selectioncolor = New_Selectioncolor
    PropertyChanged "Selectioncolor"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,0
Public Property Get SelectedItem() As Integer
    SelectedItem = m_SelectedItem
End Property

Public Property Let SelectedItem(ByVal New_SelectedItem As Integer)
    m_SelectedItem = New_SelectedItem
    PropertyChanged "SelectedItem"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=8,0,0,0
Public Property Get SelectedTextColor() As Long
    SelectedTextColor = m_SelectedTextColor
End Property

Public Property Let SelectedTextColor(ByVal New_SelectedTextColor As Long)
    m_SelectedTextColor = New_SelectedTextColor
    PropertyChanged "SelectedTextColor"
End Property

Public Sub Icons(objImagelist As Object)
    Set m_Icons = objImagelist
    Refresh
End Sub

