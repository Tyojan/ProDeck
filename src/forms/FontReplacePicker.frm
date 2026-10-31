VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.UserForm FontReplacePicker
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProDeck: All Fonts"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.CommandButton btnCancel
      Cancel          =   -1  'True
      Caption         =   "キャンセル"
      Height          =   360
      Left            =   4920
      TabIndex        =   5
      Top             =   3900
      Width           =   1320
   End
   Begin MSForms.CommandButton btnReplace
      Caption         =   "置換"
      Default         =   -1  'True
      Height          =   360
      Left            =   3480
      TabIndex        =   4
      Top             =   3900
      Width           =   1320
   End
   Begin MSForms.ComboBox cboTarget
      Height          =   315
      Left            =   300
      Style           =   0  'fmStyleDropDownCombo
      TabIndex        =   3
      Top             =   3330
      Width           =   5940
   End
   Begin MSForms.ListBox lstSource
      Height          =   1995
      Left            =   300
      TabIndex        =   1
      Top             =   945
      Width           =   5940
   End
   Begin MSForms.Label lblTarget
      Caption         =   "置換後のフォント（一覧から選択、または入力）"
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   3030
      Width           =   5400
   End
   Begin MSForms.Label lblSource
      Caption         =   "置換するフォント（このプレゼンテーションで使用中）"
      Height          =   255
      Left            =   300
      TabIndex        =   0
      Top             =   570
      Width           =   5700
   End
   Begin MSForms.Label lblHint
      Caption         =   "*（All Fonts）を選ぶと、マスター・ノート・箇条書きも含め、継承された書式を明示指定にして統一します。"
      Height          =   405
      Left            =   300
      TabIndex        =   6
      Top             =   105
      Width           =   5940
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FontReplacePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Confirmed As Boolean
Public SourceFont As String
Public TargetFont As String

Public Sub SetFontNames(ByVal fontNames As Collection)
    Dim item As Variant

    For Each item In fontNames
        If CStr(item) = "*" Then
            lstSource.AddItem "*（All Fonts）"
        Else
            lstSource.AddItem CStr(item)
            cboTarget.AddItem CStr(item)
        End If
    Next item
End Sub

Private Sub btnCancel_Click()
    Confirmed = False
    Me.Hide
End Sub

Private Sub btnReplace_Click()
    If lstSource.ListIndex < 0 Then
        MsgBox "置換するフォントを一覧から選択してください。", vbExclamation, "ProDeck: All Fonts"
        Exit Sub
    End If

    TargetFont = Trim$(cboTarget.Text)
    If Len(TargetFont) = 0 Then
        MsgBox "置換後のフォントを選択または入力してください。", vbExclamation, "ProDeck: All Fonts"
        Exit Sub
    End If

    SourceFont = CStr(lstSource.Value)
    If Left$(SourceFont, 1) = "*" Then SourceFont = "*"

    Confirmed = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Confirmed = False
End Sub
