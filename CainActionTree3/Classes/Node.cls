VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarReleative As String 'lokale Kopie
Private mvarCaption As String 'lokale Kopie
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarIcon As Integer 'lokale Kopie
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
'Private mvarChild As NodeChilds 'lokale Kopie
Public Child As NodeChilds
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarFamily As String 'lokale Kopie
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvariAppendix As Long 'lokale Kopie
Public Property Let iAppendix(ByVal vData As Long)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.iAppendix = 5
    mvariAppendix = vData
End Property


Public Property Get iAppendix() As Long
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.iAppendix
    iAppendix = mvariAppendix
End Property



Public Property Let Family(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Family = 5
    mvarFamily = vData
End Property


Public Property Get Family() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Family
    Family = mvarFamily
End Property




'Public Property Set Child(ByVal vData As NodeChilds)
''wird beim Zuweisen eines Objekts zu der Eigenschaft auf der linken Seite einer Set-Anweisung verwendet.
''Syntax: Set x.Child = Form1
'    Set mvarChild = vData
'End Property
'
'
'Public Property Get Child() As NodeChilds
''wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
''Syntax: Debug.Print X.Child
'    Set Child = mvarChild
'End Property


Public Property Let Icon(ByVal vData As Integer)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Icon = 5
    mvarIcon = vData
End Property


Public Property Get Icon() As Integer
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Icon
    Icon = mvarIcon
End Property

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



Public Property Let Releative(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Releative = 5
    mvarReleative = vData
End Property


Public Property Get Releative() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Releative
    Releative = mvarReleative
End Property



Private Sub Class_Initialize()
    Set Child = New NodeChilds
End Sub
