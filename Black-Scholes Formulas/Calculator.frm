VERSION 5.00
Begin VB.Form Calculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Option Pricing Calculator"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      Height          =   495
      Left            =   0
      Picture         =   "Calculator.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   4680
      TabIndex        =   29
      Top             =   7080
      Width           =   4740
   End
   Begin VB.PictureBox Picture4 
      Height          =   495
      Left            =   600
      Picture         =   "Calculator.frx":75C2
      ScaleHeight     =   435
      ScaleWidth      =   2700
      TabIndex        =   26
      Top             =   6390
      Width           =   2760
   End
   Begin VB.PictureBox Picture3 
      Height          =   465
      Left            =   600
      Picture         =   "Calculator.frx":B114
      ScaleHeight     =   405
      ScaleWidth      =   2355
      TabIndex        =   25
      Top             =   5880
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      Height          =   480
      Left            =   600
      Picture         =   "Calculator.frx":E4F6
      ScaleHeight     =   420
      ScaleWidth      =   1410
      TabIndex        =   24
      Top             =   5280
      Width           =   1470
   End
   Begin VB.PictureBox Picture1 
      Height          =   795
      Left            =   600
      Picture         =   "Calculator.frx":102C0
      ScaleHeight     =   735
      ScaleWidth      =   2715
      TabIndex        =   23
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox TxtCall 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox TxtD2 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox TxtD1 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox TxtVolatility 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "30.0"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox TxtRate 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "8.00"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox TxtYearstoMature 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "0.25"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox TxtStrikePrice 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "65.00"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox TxtStockPrice 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "60.00"
      Top             =   840
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Black-Scholes"
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3375
      Begin VB.TextBox TxtPut 
         Height          =   285
         Left            =   1320
         TabIndex        =   30
         Top             =   3285
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Put:"
         Height          =   255
         Left            =   1020
         TabIndex        =   31
         Top             =   3330
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "(d2)"
         Height          =   255
         Left            =   2970
         TabIndex        =   22
         Top             =   2580
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "(d1)"
         Height          =   255
         Left            =   2970
         TabIndex        =   21
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "(v)"
         Height          =   255
         Left            =   3000
         TabIndex        =   20
         Top             =   1725
         Width           =   255
      End
      Begin VB.Label Label15 
         Caption         =   "(r)"
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   1470
         Width           =   255
      End
      Begin VB.Label Label14 
         Caption         =   "(T)"
         Height          =   255
         Left            =   3000
         TabIndex        =   18
         Top             =   1230
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "(X)"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   990
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "(S)"
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   750
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Stock price:"
         Height          =   255
         Left            =   450
         TabIndex        =   14
         Top             =   750
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Strike price:"
         Height          =   255
         Left            =   465
         TabIndex        =   13
         Top             =   990
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Years to maturity:"
         Height          =   255
         Left            =   90
         TabIndex        =   12
         Top             =   1230
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Rate (continuos):"
         Height          =   255
         Left            =   90
         TabIndex        =   11
         Top             =   1470
         Width           =   1230
      End
      Begin VB.Label Label5 
         Caption         =   "Volatility:"
         Height          =   255
         Left            =   690
         TabIndex        =   10
         Top             =   1725
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Call:"
         Height          =   255
         Left            =   1005
         TabIndex        =   9
         Top             =   3060
         Width           =   375
      End
   End
   Begin VB.Label Label17 
      Caption         =   "(Put)"
      Height          =   255
      Left            =   225
      TabIndex        =   28
      Top             =   6525
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "(Call)"
      Height          =   255
      Left            =   210
      TabIndex        =   27
      Top             =   6015
      Width           =   375
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Black-Scholes (1973) Option Pricing Formula
'Original code by Charles Toepfer: toepfer_c@hotmail.com
'Please acknowledge use of this code by including this header.

Const Pi = 3.14159265358979
'3.141592653589793238462643


'This is a numerical approximation to the normal distribution:
    Const a1 = 0.31938153
    Const a2 = -0.356563782
    Const a3 = 1.781477937:
    Const a4 = -1.821255978
    Const a5 = 1.330274429
'------------------------------------

Sub GBlackScholes(S As Double, X As Double, T As Double, r As Double, v As Double)
  
Dim d1 As Double, d2 As Double
        
  'Fix Percentages:
  v = (v * (0.01))
  r = (r * (0.01))
   
  d1 = (Math.Log(S / X) + (r + v ^ 2 / 2) * T) / ((v) * Math.Sqr(T))
  d2 = d1 - v * Math.Sqr(T)
    
  'Show Values:
  TxtD1.Text = CStr(d1)
  TxtD2.Text = CStr(d2)
    
  'Call
  TxtCall.Text = (S * CND(d1) - X * Math.Exp(-r * T) * CND(d2))
  'Put
  TxtPut.Text = (X * Math.Exp(-r * T) * CND(-d2) - S * CND(-d1))
 
End Sub

'The cumulative normal distribution function:
Public Function CND(X As Double) As Double

    Dim L As Double, K As Double

    L = Math.Abs(X)
    K = 1 / (1 + 0.2316419 * L)
    CND = 1 - 1 / Math.Sqr(2 * Pi) * Math.Exp(-L ^ 2 / 2) * (a1 * K + a2 * K ^ 2 + a3 * K ^ 3 + a4 * K ^ 4 + a5 * K ^ 5)

    If X < 0 Then
      CND = 1 - CND
    End If

End Function

Private Sub Command1_Click()

     Call GBlackScholes(TxtStockPrice.Text, TxtStrikePrice.Text, TxtYearstoMature.Text, TxtRate.Text, TxtVolatility.Text)

End Sub

Private Sub Label18_Click()

End Sub

