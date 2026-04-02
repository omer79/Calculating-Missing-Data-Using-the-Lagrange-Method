VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Form1 
   Caption         =   "lagrange ÈÇÓÊÎÏÇã ØÑíÞÉ missing data ÈÑäÇãÌ áÍÓÇÈ"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   9360
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.TextBox Text16 
      Height          =   495
      Left            =   1560
      TabIndex        =   25
      Text            =   "x"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Íæá ÇáÈÑäÇãÌ"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÎÑæÌ"
      Height          =   495
      Left            =   4200
      TabIndex        =   16
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÇÍÓÈ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   15
      Top             =   3600
      Width           =   6135
   End
   Begin VB.TextBox Text15 
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Text            =   "0"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   13
      Text            =   "0"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox Text13 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Text            =   "0"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Text            =   "0"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   405
      Left            =   1800
      TabIndex        =   10
      Text            =   "0"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   7320
      TabIndex        =   9
      Text            =   "0"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      Text            =   "0"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Text            =   "0"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Text            =   "0"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Text            =   "0"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   7320
      TabIndex        =   4
      Text            =   "0"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Text            =   "0"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Text            =   "0"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Text            =   "0"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Text            =   "0"
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   24
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "f(x)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   23
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "f(x1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   22
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "p1(x1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   21
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "p2(x1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   20
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "p3(x1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   19
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "p4(x1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   18
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function func_f_x(x1)
ScriptControl1.Reset
ScriptControl1.ExecuteStatement ("x=" & x1)
func_f_x = ScriptControl1.Eval(Trim(Text16.Text))
End Function
Private Sub Command1_Click()
Dim x1
If Text6.Text = 0 Then
Text6.Text = func_f_x(Val(Text1.Text))
End If
If Text7.Text = 0 Then
Text7.Text = func_f_x(Val(Text2.Text))
End If
If Text8.Text = 0 Then
Text8.Text = func_f_x(Val(Text3.Text))
End If
If Text9.Text = 0 Then
Text9.Text = func_f_x(Val(Text4.Text))
End If
If Text10.Text = 0 Then
Text10.Text = func_f_x(Val(Text5.Text))
End If
 Dim v, a0, a1, a2, a3, a4, b0, b1, b2, b3, b4, t1, t2, t3, t4, t5, t6, t7, t8, t9, t10 As Double
        t1 = Val(Text1.Text)
        t2 = Val(Text2.Text)
        t3 = Val(Text3.Text)
        t4 = Val(Text4.Text)
        t5 = Val(Text5.Text)
        t6 = Val(Text6.Text)
        t7 = Val(Text7.Text)
        t8 = Val(Text8.Text)
        t9 = Val(Text9.Text)
        t10 = Val(Text10.Text)

        v = Val(Text15.Text)
        If t1 > t2 Or t2 > t3 Or t3 > t4 Or t4 > t5 Then
        MsgBox "backward ÇãÇ Çäß ÇÚØíÊ äÞØÉ ÈÕÝÑ Çæ Çäß ÊÓÊÎÏã ØÑíÞÉ forward  äÍä äÓÊÎÏã ØÑíÞÉ "
        End If

        If v > t1 And v < t2 Then
            a0 = t1
            a1 = t2
            a2 = t3
            a3 = t4
            a4 = t5
            b0 = t6
            b1 = t7
            b2 = t8
            b3 = t9
            b4 = t10
        End If
        If v > t2 And v < t3 Then
            a0 = t2
            a1 = t3
            b0 = t7
            b1 = t8
            a4 = t5
            b4 = t10
            a2 = t4
            b2 = t9
            a3 = t1
            b3 = t6

        End If
        If v > t3 And v < t4 Then
            a0 = t3
            a1 = t4
            b0 = t8
            b1 = t9
            a4 = t1
            b4 = t6
            a2 = t2
            a3 = t5
            b2 = t7
            b3 = t10
        End If
          
        If v > t4 And v < t5 Then
            a0 = t4
            a1 = t5
            a2 = t3
            a3 = t2
            a4 = t1
            b0 = t9
            b1 = t10
            b2 = t8
            b3 = t7
            b4 = t6
        End If
        If v >= t5 Or v <= t1 Then
         MsgBox "íÌÈ Çä Êßæä ÇáäÞÉ ÇáãØáæÈÉ ÇÞá ãä ÎÇãÓ äÞØÉ æÇßÈÑ ãä Çæá äÞØÉ"
         Text11.Text = 0
         Text11.Text = 0
         Text11.Text = 0
         Text11.Text = 0
        Else
        Text11.Text = (b0 * (v - a1)) / (a0 - a1) + (b1 * (v - a0)) / (a1 - a0)
        Text12.Text = (b0 * (v - a1) * (v - a2)) / ((a0 - a2) * (a0 - a1)) + (b1 * (v - a0) * (v - a2)) / ((a1 - a0) * (a1 - a2)) + (b2 * (v - a0) * (v - a1)) / ((a2 - a0) * (a2 - a1))
        Text13.Text = (b0 * (v - a1) * (v - a2) * (v - a3)) / ((a0 - a2) * (a0 - a1) * (a0 - a3)) + (b1 * (v - a0) * (v - a2) * (v - a3)) / ((a1 - a0) * (a1 - a2) * (a1 - a3)) + (b2 * (v - a0) * (v - a1) * (v - a3)) / ((a2 - a0) * (a2 - a1) * (a2 - a3)) + (b3 * (v - a0) * (v - a1) * (v - a2)) / ((a3 - a0) * (a3 - a1) * (a3 - a2))
        Text14.Text = (b0 * (v - a1) * (v - a2) * (v - a3) * (v - a4)) / ((a0 - a2) * (a0 - a1) * (a0 - a3) * (a0 - a4)) + (b1 * (v - a0) * (v - a2) * (v - a3) * (v - a4)) / ((a1 - a0) * (a1 - a2) * (a1 - a3) * (a1 - a4)) + (b2 * (v - a0) * (v - a1) * (v - a3) * (v - a4)) / ((a2 - a0) * (a2 - a1) * (a2 - a3) * (a2 - a4)) + (b3 * (v - a0) * (v - a1) * (v - a2) * (v - a4)) / ((a3 - a0) * (a3 - a1) * (a3 - a2) * (a3 - a4)) + (b4 * (v - a0) * (v - a1) * (v - a2) * (v - a3)) / ((a4 - a0) * (a4 - a1) * (a4 - a2) * (a4 - a3))
        End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form2.Show
Me.Hide
End Sub

Private Sub Text1_Change()
 If Not IsNumeric(Text1.Text) Then
            Text1.Text = "0"
        End If
        If Text1.Text = "" Then
            Text1.Text = "0"
        End If
End Sub

Private Sub Text10_Change()
If Not IsNumeric(Text10.Text) Then
            Text10.Text = "0"
        End If
        If Text10.Text = "" Then
            Text10.Text = "0"
        End If
End Sub

Private Sub Text15_Change()
If Not IsNumeric(Text15.Text) Then
            Text15.Text = "0"
        End If
        If Text15.Text = "" Then
            Text15.Text = "0"
        End If
End Sub

Private Sub Text2_Change()
If Not IsNumeric(Text2.Text) Then
            Text2.Text = "0"
        End If
        If Text2.Text = "" Then
            Text2.Text = "0"
        End If
End Sub

Private Sub Text3_Change()
If Not IsNumeric(Text3.Text) Then
            Text3.Text = "0"
        End If
        If Text3.Text = "" Then
            Text3.Text = "0"
        End If
End Sub

Private Sub Text4_Change()
If Not IsNumeric(Text4.Text) Then
            Text4.Text = "0"
        End If
        If Text4.Text = "" Then
            Text4.Text = "0"
        End If
End Sub

Private Sub Text5_Change()
If Not IsNumeric(Text5.Text) Then
            Text5.Text = "0"
        End If
        If Text5.Text = "" Then
            Text5.Text = "0"
        End If
End Sub

Private Sub Text6_Change()
If Not IsNumeric(Text6.Text) Then
            Text6.Text = "0"
        End If
        If Text6.Text = "" Then
            Text6.Text = "0"
        End If
End Sub

Private Sub Text7_Change()
If Not IsNumeric(Text7.Text) Then
            Text7.Text = "0"
        End If
        If Text7.Text = "" Then
            Text7.Text = "0"
        End If
End Sub

Private Sub Text8_Change()
If Not IsNumeric(Text8.Text) Then
            Text8.Text = "0"
        End If
        If Text8.Text = "" Then
            Text8.Text = "0"
        End If
End Sub

Private Sub Text9_Change()
If Not IsNumeric(Text9.Text) Then
            Text9.Text = "0"
        End If
        If Text9.Text = "" Then
            Text9.Text = "0"
        End If
End Sub
