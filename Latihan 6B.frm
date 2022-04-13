VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Selesai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   24
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox Text7 
      Height          =   405
      Left            =   3480
      TabIndex        =   23
      Top             =   6480
      Width           =   3735
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3480
      TabIndex        =   22
      Top             =   5880
      Width           =   3735
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   5280
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   3480
      TabIndex        =   20
      Top             =   4680
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   3480
      TabIndex        =   19
      Top             =   4080
      Width           =   3735
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   9960
      TabIndex        =   13
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Beri tanda jika kawin"
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Jenis Kelamin"
      Height          =   975
      Left            =   7320
      TabIndex        =   7
      Top             =   2400
      Width           =   4815
      Begin VB.OptionButton Option2 
         Caption         =   "Perempuan"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Laki-laki"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3480
      TabIndex        =   6
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   3480
      TabIndex        =   5
      Top             =   1800
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label11 
      Caption         =   "Gaji Bersih"
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label Label10 
      Caption         =   "Pajak"
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Label Label9 
      Caption         =   "Tunjangan Anak"
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Tunjangan Kawin"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "Gaji Pokok"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Jumlah Anak"
      Height          =   255
      Left            =   7320
      TabIndex        =   12
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Status Perkawinan"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Bagian"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Pegawai"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Nomor Pegawai"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Formulir Perhitungan Gaji Pegawai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
 If Check1.Value = 1 Then
    Text4.Text = Val(Text3.Text) * 0.1
    If Val(Combo2.Text) >= (3 And Check1.Value = 1) Then
        Text5.Text = 3 * 0.1 * Val(Text3.Text)
    Else
        Text5.Text = Val(Combo2.Text) * 0.1 * Val(Text3.Text)
    End If
Else
Text4.Text = 0
Text5.Text = 0
End If
'MENGHITUNG GAJI POKOK
 If Combo1.Text = "Akuntansi" Then
        Text3 = 750000
    Else
        If Combo1.Text = "Administrasi Umum" Then
            Text3 = 500000
        Else
            If Combo1.Text = "Produksi" Then
                Text3 = 600000
            Else
                Text3 = 500000
        End If
    End If
End If
'MENGHITUNG PAJAK DAN TOTAL GAJI
Text6.Text = (Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text)) * 0.15
Text7.Text = (Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text)) - Val(Text6.Text)
End Sub

Private Sub Combo1_Click()
'MENGHITUNG GAJI POKOK
 If Combo1.Text = "Akuntansi" Then
        Text3 = 750000
    Else
        If Combo1.Text = "Administrasi Umum" Then
            Text3 = 500000
        Else
            If Combo1.Text = "Produksi" Then
                Text3 = 600000
            Else
                Text3 = 500000
        End If
    End If
End If
'MENGHITUNG TUNJANGAN KAWIN DA TUNJANGAN ANAK
If Check1.Value = 1 Then
    Text4.Text = Val(Text3.Text) * 0.1
    If Val(Combo2.Text) >= (3 And Check1.Value = 1) Then
        Text5.Text = 3 * 0.1 * Val(Text3.Text)
    Else
        Text5.Text = Val(Combo2.Text) * 0.1 * Val(Text3.Text)
    End If
Else
    Text4.Text = 0
    Text5.Text = 0
End If

'MENGHITUNG PAJAK DAN TOTAL GAJI
Text6.Text = (Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text)) * 0.15
Text7.Text = (Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text)) - Val(Text6.Text)

End Sub
Private Sub Combo2_Click()
If Check1.Value = 1 Then
    If Combo2.Text >= 3 Then
        Text5.Text = 3 * 0.1 * Val(Text3.Text)
    Else
        Text5.Text = Val(Combo2.Text) * 0.1 * Val(Text3.Text)
    End If
Else
    Text5.Text = 0
End If
'MENGHITUNG GAJI POKOK
 If Combo1.Text = "Akuntansi" Then
        Text3 = 750000
    Else
        If Combo1.Text = "Administrasi Umum" Then
            Text3 = 500000
        Else
            If Combo1.Text = "Produksi" Then
                Text3 = 600000
            Else
                Text3 = 500000
        End If
    End If
End If
'MENGHITUNG PAJAK DAN TOTAL GAJI
Text6.Text = (Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text)) * 0.15
Text7.Text = (Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text)) - Val(Text6.Text)
End Sub

Private Sub Command1_Click()
 End
End Sub

Private Sub Form_Activate()
 Combo1.AddItem "Akuntansi"
 Combo1.AddItem "Administrasi Umum"
 Combo1.AddItem "Produksi"
 Combo1.AddItem "Pengamanan"
 Combo2.AddItem 0
 Combo2.AddItem 1
 Combo2.AddItem 2
 Combo2.AddItem 3
 Combo2.AddItem 4
 Combo2.AddItem 5
 Combo2.AddItem 6
 Combo2.AddItem 7
 Combo2.AddItem 8
 Combo2.AddItem 9
 Combo2.AddItem 10
End Sub

Private Sub Text5_LostFocus()
 Text6.Text = (Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text)) * 0.15
 Text7.SetFocus
End Sub

Private Sub Text6_LostFocus()
 Text7.Text = (Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text)) - Val(Text6.Text)
 Command1.SetFocus
End Sub
