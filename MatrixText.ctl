VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.UserControl MatrixText 
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2805
   ScaleHeight     =   1830
   ScaleWidth      =   2805
   Begin SHDocVwCtl.WebBrowser WB1 
      Height          =   1680
      Left            =   -45
      TabIndex        =   1
      Top             =   0
      Width           =   2760
      ExtentX         =   4868
      ExtentY         =   2963
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   6570
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "MatrixText.ctx":0000
      Top             =   5580
      Width           =   510
   End
End
Attribute VB_Name = "MatrixText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Default Property Values:
Const m_def_MatrixText = "Matrix Text"

'Property Variables:
Dim m_MatrixText As String



Private Sub UserControl_Resize()

  WB1.Move -100, -250, _
           (Width + 400), _
           (Height + 300)
 
End Sub

Private Sub UserControl_Show()

  WB1.Navigate "about:blank"

End Sub

Private Sub WB1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
  
  DoEvents
  DoEvents
  
  WB1.Document.write Text1 & vbCrLf & _
    "<body oncontextmenu='return false' bgcolor='000000'>" & vbCrLf & _
    "<div id='matrix'>" & m_MatrixText & "</div>" & vbCrLf & _
    "</body>" & vbCrLf & _
    "</html>"
    
  WB1.Refresh
  
End Sub
 
'matrixtext
Public Property Get MatrixText() As String
    MatrixText = m_MatrixText
End Property
Public Property Let MatrixText(ByVal New_MatrixText As String)
    m_MatrixText = New_MatrixText
    PropertyChanged "MatrixText"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_MatrixText = m_def_MatrixText
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_MatrixText = PropBag.ReadProperty("MatrixText", m_def_MatrixText)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("MatrixText", m_MatrixText, m_def_MatrixText)
End Sub

 
