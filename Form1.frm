VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menampilkan Font yang Terinstall di PC Anda"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2040
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ketika Anda mengklik salah satu font, jenis huruf pada 'listbox akan berubah sesuai dengan jenis huruf yang 'diklik saat itu. Nama font yang diklik akan 'ditampilkan di textbox.

Private Sub Form_Load() 'Tampilkan semua font ke dalam 'ListBox
Dim counter As Integer
  For counter = 0 To Screen.FontCount - 1
      List1.AddItem Screen.Fonts(counter)
  Next
End Sub

Private Sub List1_Click() 'Jika salah satu font 'diklik...
Static tempheight As Single
  If tempheight = 0 Then tempheight = List1.Height
  Text1.Text = List1.List(List1.ListIndex)
  List1.Font.Name = List1.List(List1.ListIndex)
  List1.Height = tempheight
End Sub

