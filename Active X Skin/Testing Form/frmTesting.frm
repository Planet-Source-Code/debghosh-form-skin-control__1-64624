VERSION 5.00
Object = "*\A..\Skin Ctl\Deba Skin Control.vbp"
Begin VB.Form frmTesting 
   Caption         =   "Testing Form"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10485
   Icon            =   "frmTesting.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin SkinCtl.sK sK1 
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      CaptionColor    =   16761024
      FormBackColor   =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SkinCtl.GrdBtn GrdBtn3 
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   3720
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      DefaultGradient =   1
      Caption         =   "Check This Button"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12648447
      ScaleHeight     =   41
      ScaleMode       =   3
      ScaleWidth      =   313
      OnMouseMoveForeColor=   16777215
   End
   Begin SkinCtl.GrdBtn GrdBtn2 
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   2760
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      DefaultGradient =   3
      OnMouseMoveGradient=   1
      Caption         =   "Check This Button"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleHeight     =   41
      ScaleMode       =   3
      ScaleWidth      =   313
      OnMouseMoveForeColor=   8454143
   End
   Begin SkinCtl.GrdBtn GrdBtn1 
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   1800
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      Caption         =   "Check This Button"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      ScaleHeight     =   41
      ScaleMode       =   3
      ScaleWidth      =   313
      OnMouseMoveForeColor=   16777215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move your mouse over this Buttons and check gradient effect"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1920
      TabIndex        =   3
      Top             =   5040
      Width           =   6690
   End
End
Attribute VB_Name = "frmTesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
