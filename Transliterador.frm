VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tranliterar"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   Icon            =   "Transliterador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Toranik"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   420
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   4740
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   420
      TabIndex        =   0
      Top             =   630
      Width           =   4740
   End
   Begin VB.Label Label2 
      Caption         =   "Español:"
      Height          =   225
      Left            =   420
      TabIndex        =   3
      Top             =   210
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "Transliteración hebrea:"
      Height          =   225
      Left            =   420
      TabIndex        =   2
      Top             =   1260
      Width           =   1905
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

On Error Resume Next

Dim s As String
Dim A() As String
Dim B() As String
s = Text1.Text
s = Replace$(s, "  ", " ")
s = Replace$(s, "  ", " ")
s = Replace$(s, "  ", " ")

A = Split(s, " ")
Dim i As Long
Dim j As Long
ReDim B(UBound(A)) As String

For i = 0 To UBound(A)
j = UBound(A) - i
    B(j) = TkTranslit(A(i))
Next i

Text2.Text = Join(B, " ")

End Sub
