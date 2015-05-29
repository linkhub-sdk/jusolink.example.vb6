VERSION 5.00
Begin VB.Form jusolinkExam 
   BorderStyle     =   1  '���� ����
   Caption         =   "�ּҸ�ũ API SDK ����"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9030
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   9030
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame2 
      Caption         =   "���� API"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   10
      Top             =   2280
      Width           =   4695
      Begin VB.CommandButton btnGetBalance 
         Caption         =   "�ܿ�����Ʈ ��ȸ"
         Height          =   495
         Left            =   2420
         TabIndex        =   6
         Top             =   300
         Width           =   1935
      End
      Begin VB.CommandButton btnUnitCost 
         Caption         =   "�˻��ܰ� ��ȸ"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   300
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�ּҰ˻� API"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   7815
      Begin VB.TextBox txtAddress 
         Height          =   270
         Left            =   1320
         TabIndex        =   3
         Top             =   835
         Width           =   4815
      End
      Begin VB.TextBox txtSectionNum 
         Height          =   270
         Left            =   4440
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtZipcode 
         Height          =   270
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton btnSearchForm 
         Caption         =   "�ּҰ˻�"
         Height          =   855
         Left            =   6360
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "�������ȣ :"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "���ּ� :"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "�����ȣ :"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "jusolinkExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��ũ���̵�
Private Const linkID = "TESTER_JUSO"
'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "FjaRgAfVUPvSDHTrdd/uw/dt/Cdo3GgSFKyE1+NQ+bc="

Private JusolinkService As New Jusolink

Public zipcode As String
Public sectionNum As String
Public detailAddress As String
'�ܿ�����Ʈ ��ȸ
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = JusolinkService.GetBalance()
    
    If balance < 0 Then
        MsgBox ("[" + CStr(JusolinkService.LastErrCode) + "] " + JusolinkService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub
'�˻��ܰ� Ȯ��
Private Sub btnUnitCost_Click()
    Dim unitCost As Single
    
    unitCost = JusolinkService.GetUnitCost()
    
    If unitCost < 0 Then
        MsgBox ("[" + CStr(JusolinkService.LastErrCode) + "] " + JusolinkService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�˻� �ܰ� : " + CStr(unitCost)
End Sub
'�ּҰ˻��� ȣ��
Private Sub btnSearchForm_Click()
     searchForm.linkID = linkID
     searchForm.SecretKey = SecretKey
     searchForm.Show
End Sub
Private Sub Form_Activate()
    txtZipcode.Text = zipcode
    txtSectionNum.Text = sectionNum
    txtAddress.Text = detailAddress
    txtAddress.SetFocus
    
    If Len(txtAddress.Text) > 0 Then
        txtAddress.SelStart = Len(txtAddress.Text)
    End If
End Sub
Private Sub Form_Load()
    JusolinkService.Initialize linkID, SecretKey
End Sub


