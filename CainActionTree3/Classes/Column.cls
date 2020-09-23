VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Column"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarColWidth As Integer 'lokale Kopie
Private mvarCaption As String 'lokale Kopie
Public Property Let Caption(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Caption = 5
    mvarCaption = vData
End Property


Public Property Get Caption() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
End Property



Public Property Let ColWidth(ByVal vData As Integer)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.ColWidth = 5
    mvarColWidth = vData
End Property


Public Property Get ColWidth() As Integer
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.ColWidth
    ColWidth = mvarColWidth
End Property



