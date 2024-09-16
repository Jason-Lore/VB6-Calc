VERSION 5.00
Begin VB.Form Calculator 
   Caption         =   "Calculator"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   3555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_nums 
      Caption         =   "3"
      Height          =   615
      Index           =   3
      Left            =   1800
      TabIndex        =   20
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton btn_func_abs 
      Caption         =   "+/-"
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton btn_nums 
      Caption         =   "."
      Height          =   615
      Index           =   10
      Left            =   1800
      TabIndex        =   18
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton btn_op_equal 
      Caption         =   "="
      Height          =   615
      Index           =   3
      Left            =   2640
      TabIndex        =   17
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton btn_op 
      Caption         =   "%"
      Height          =   615
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton btn_func_clear 
      Caption         =   "CE"
      Height          =   615
      Left            =   960
      TabIndex        =   15
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton btn_func_clear_all 
      Caption         =   "C"
      Height          =   615
      Left            =   1800
      TabIndex        =   14
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton btn_op 
      Caption         =   "/"
      Height          =   615
      Index           =   3
      Left            =   2640
      TabIndex        =   13
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton btn_op 
      Caption         =   "*"
      Height          =   615
      Index           =   0
      Left            =   2640
      TabIndex        =   12
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton btn_op 
      Caption         =   "-"
      Height          =   615
      Index           =   1
      Left            =   2640
      TabIndex        =   11
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton btn_op 
      Caption         =   "+"
      Height          =   615
      Index           =   2
      Left            =   2640
      TabIndex        =   10
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton btn_nums 
      Caption         =   "9"
      Height          =   615
      Index           =   9
      Left            =   1800
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton btn_nums 
      Caption         =   "8"
      Height          =   615
      Index           =   8
      Left            =   960
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton btn_nums 
      Caption         =   "7"
      Height          =   615
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton btn_nums 
      Caption         =   "6"
      Height          =   615
      Index           =   6
      Left            =   1800
      TabIndex        =   6
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton btn_nums 
      Caption         =   "5"
      Height          =   615
      Index           =   5
      Left            =   960
      TabIndex        =   5
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton btn_nums 
      Caption         =   "4"
      Height          =   615
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton btn_nums 
      Caption         =   "1"
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton btn_nums 
      Caption         =   "0"
      Height          =   615
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton btn_nums 
      Caption         =   "2"
      Height          =   615
      Index           =   2
      Left            =   960
      TabIndex        =   1
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txt_display 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentInput As String
Dim FirstOperand As Double
Dim Operator As String

Private Sub btn_func_abs_Click()
    txt_display.Text = Abs(txt_display.Text)
End Sub

Private Sub btn_func_clear_all_Click()
    CurrentInput = ""
    Operator = ""
    FirstOperand = 0
    txt_display.Text = 0

End Sub

Private Sub btn_func_clear_Click()
    CurrentInput = ""
    txt_display.Text = 0
End Sub

Private Sub btn_nums_Click(Index As Integer)
    CurrentInput = CurrentInput & btn_nums(Index).Caption
    txt_display.Text = CurrentInput
End Sub

Private Sub btn_op_Click(Index As Integer)
    FirstOperand = Val(CurrentInput)
    
    Operator = btn_op(Index).Caption
    
    CurrentInput = ""
End Sub

Private Sub btn_op_equal_Click(Index As Integer)
    Dim SecondOperand As Double
    Dim Result As Double
    
    SecondOperand = Val(CurrentInput)
    
    Select Case Operator
        Case "+"
            Result = FirstOperand + SecondOperand
        Case "-"
            Result = FirstOperand - SecondOperand
        Case "*"
            Result = FirstOperand * SecondOperand
        Case "%"
            Result = FirstOperand Mod SecondOperand
        Case "/"
            If SecondOperand <> 0 Then
                Result = FirstOperand / SecondOperand
            Else
                MsgBox ("Cannot Divide by 0")
                Exit Sub
            End If
    End Select
    
    txt_display.Text = Result
    CurrentInput = ""
    Operator = ""
    FirstOperand = Result
        
End Sub
