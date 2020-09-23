VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Style Controls"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   Icon            =   "New Style Controls.frx":0000
   LinkTopic       =   "New Style Controls"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3075
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   1800
      Width           =   2775
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   3075
      TabIndex        =   2
      Top             =   375
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   150
      TabIndex        =   1
      Top             =   375
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Old ComboBox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3075
      TabIndex        =   7
      Top             =   1575
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "New ComboBox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   150
      TabIndex        =   6
      Top             =   1575
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Old ListBox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3075
      TabIndex        =   5
      Top             =   150
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "New ListBox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   150
      TabIndex        =   4
      Top             =   150
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This Source Code is Made By:
' Fernando Macedo
' E-mail: fjmm007@hotmail.com
' 16 september 2007

' --- New Style Controls ---
' A new way to present ComboBox and ListBox

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)


Private Sub Form_Load()

Dim m_Style_Cmb As Long
' Get the current EXSTYLE attributes for the Combo
m_Style_Cmb = GetWindowLong(Combo1.hwnd, GWL_EXSTYLE)
' Modify the EXSTYLE to show
' Put the text in right position
m_Style_Cmb = m_Style_Cmb Or &H3000
Call SetWindowLong(Combo1.hwnd, GWL_EXSTYLE, m_Style_Cmb)

Combo1.AddItem "New_Style_Controls_0"
Combo1.AddItem "New_Style_Controls_1"
Combo1.AddItem "New_Style_Controls_2"
Combo1.AddItem "New_Style_Controls_3"
Combo1.AddItem "New_Style_Controls_4"
Combo1.AddItem "New_Style_Controls_5"
Combo1.AddItem "New_Style_Controls_6"
Combo1.AddItem "New_Style_Controls_7"
Combo1.AddItem "New_Style_Controls_8"
Combo1.ListIndex = 0

Combo2.AddItem "Old_Style_Controls_0"
Combo2.AddItem "Old_Style_Controls_1"
Combo2.AddItem "Old_Style_Controls_2"
Combo2.AddItem "Old_Style_Controls_3"
Combo2.AddItem "Old_Style_Controls_4"
Combo2.AddItem "Old_Style_Controls_5"
Combo2.AddItem "Old_Style_Controls_6"
Combo2.AddItem "Old_Style_Controls_7"
Combo2.AddItem "Old_Style_Controls_8"
Combo2.ListIndex = 0

List1.AddItem "New_Style_Controls_0"
List1.AddItem "New_Style_Controls_1"
List1.AddItem "New_Style_Controls_2"
List1.AddItem "New_Style_Controls_3"
List1.AddItem "New_Style_Controls_4"
List1.AddItem "New_Style_Controls_5"
List1.AddItem "New_Style_Controls_6"
List1.AddItem "New_Style_Controls_7"
List1.AddItem "New_Style_Controls_8"

List2.AddItem "Old_Style_Controls_0"
List2.AddItem "Old_Style_Controls_1"
List2.AddItem "Old_Style_Controls_2"
List2.AddItem "Old_Style_Controls_3"
List2.AddItem "Old_Style_Controls_4"
List2.AddItem "Old_Style_Controls_5"
List2.AddItem "Old_Style_Controls_6"
List2.AddItem "Old_Style_Controls_7"
List2.AddItem "Old_Style_Controls_8"

Dim m_Style_Lst As Long
' Get the current EXSTYLE attributes for the Combo
m_Style_Lst = GetWindowLong(List1.hwnd, GWL_EXSTYLE)
' Modify the EXSTYLE to show
' Put the text in right position
' m_Style_Lst = m_Style Or &H4000 ( Text Left )
' m_Style_Lst = m_Style Or &H5000 ( Text Right )
m_Style_Lst = m_Style_Lst Or &H5000
Call SetWindowLong(List1.hwnd, GWL_EXSTYLE, m_Style_Lst)


End Sub


