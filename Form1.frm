VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menonaktifkan Tombol Close di Form"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   1680
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hilangkan Tombol Close di Form"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Jadi, form ditutup harus melalui Command1
'Unload me tidak berfungsi di sini, jadi kita 'menggunakan End, bebaskan memory sebelumnya
Private Sub Command1_Click()
  'Unload Me
   RemoveCancelMenuItem Me  'Hilangkan tombol Close di
                           'form ini
End Sub

Private Sub Command2_Click()
    End
End Sub
