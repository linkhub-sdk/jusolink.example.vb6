VERSION 5.00
Begin VB.Form detailForm 
   BorderStyle     =   1  '���� ����
   Caption         =   "���ּ� �Է�"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   13860
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton btnResearch 
      Caption         =   "�ٽð˻�"
      Height          =   495
      Left            =   6960
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ȯ��"
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Index           =   1
      Left            =   7080
      TabIndex        =   9
      Top             =   1800
      Width           =   6495
      Begin VB.TextBox txtJibunDetail 
         Height          =   300
         Left            =   1320
         TabIndex        =   4
         Top             =   1125
         Width           =   4935
      End
      Begin VB.Label txtJibunAddr 
         Height          =   495
         Left            =   1320
         TabIndex        =   23
         Top             =   585
         Width           =   4695
      End
      Begin VB.Label txtJibunSectionNum 
         Height          =   255
         Left            =   4680
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label txtJibunZipcode 
         Height          =   210
         Left            =   1560
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "���ּ� : "
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "�⺻�ּ� :"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "�������ȣ :"
         Height          =   255
         Left            =   3480
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "�������ȣ : "
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.OptionButton optJibun 
      Caption         =   "ǥ��ȭ �����ּ�"
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   6495
      Begin VB.TextBox txtRoadDetail 
         Height          =   300
         Left            =   1320
         TabIndex        =   2
         Top             =   1125
         Width           =   4935
      End
      Begin VB.Label txtRoadAddr 
         Height          =   495
         Left            =   1320
         TabIndex        =   22
         Top             =   555
         Width           =   4455
      End
      Begin VB.Label txtRoadSectionNum 
         Height          =   255
         Left            =   4440
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label txtRoadZipcode 
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "�������ȣ : "
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "���ּ� :"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "�⺻�ּ� : "
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "�������ȣ :"
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.OptionButton optRoad 
      Caption         =   "ǥ��ȭ ���θ��ּ� "
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "> ��Ȯ�� ���� �߼��� ���� ǥ��ȭ ���θ� �ּ� ����� �����մϴ�."
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   795
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "> �Ʒ��� �ּҸ� Ȯ���Ͻð� �����Ͻ� �� Ȯ�ι�ư�� ��������."
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5535
   End
End
Attribute VB_Name = "detailForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public zipcode As String
Public sectionNum As String
Public roadAddr1 As String
Public roadAddr2 As String
Public jibunAddr As String
Public addrType
'Ȯ�ι�ư
Private Sub btnOk_Click()
    jusolinkExam.zipcode = zipcode    '�����ȣ
    jusolinkExam.sectionNum = sectionNum '�������ȣ
        
    If optRoad.value = True Then
        'ǥ��ȭ ���θ��ּ�
        jusolinkExam.detailAddress = roadAddr1 + ", " + txtRoadDetail.Text + " " + roadAddr2
    Else
        'ǥ��ȭ �����ּ�
        jusolinkExam.detailAddress = jibunAddr + ", " + txtJibunDetail.Text
    End If
    
    Unload Me
    jusolinkExam.Show
End Sub
'�ٽð˻�
Private Sub btnResearch_Click()
    Load searchForm
    searchForm.Show
    searchForm.txtIndex.SetFocus
    searchForm.ListView1.ListItems.Clear
    Unload Me
End Sub
Private Sub optJibun_Click()
    txtJibunDetail.SetFocus
End Sub
Private Sub optRoad_Click()
    txtRoadDetail.SetFocus
End Sub
Private Sub txtRoadDetail_GotFocus()
    optRoad.value = True
End Sub
Private Sub txtJibunDetail_GotFocus()
    optJibun.value = True
End Sub
Private Sub txtRoadDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call btnOk_Click
    End If
End Sub
Private Sub txtJibunDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call btnOk_Click
    End If
End Sub
Private Sub Form_Load()
    txtRoadZipcode.Caption = zipcode
    txtRoadSectionNum.Caption = sectionNum
    txtRoadAddr.Caption = roadAddr1 + " " + roadAddr2
    
    txtJibunZipcode.Caption = zipcode
    txtJibunSectionNum.Caption = sectionNum
    txtJibunAddr.Caption = jibunAddr
End Sub

