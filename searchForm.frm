VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form searchForm 
   BorderStyle     =   1  '���� ����
   Caption         =   " �ּҰ˻� Example"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   17970
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton btnPrevPage 
      Caption         =   "���� ������"
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton btnNextPage 
      Caption         =   "���� ������"
      Height          =   375
      Left            =   9600
      TabIndex        =   5
      Top             =   8640
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   16680
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7095
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   12515
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�����ȣ"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�������ȣ"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "���θ��ּ�"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�ΰ�����"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "�����ּ�"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "�����������"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "�ּҰ˻�"
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin VB.CommandButton btnSearch 
         Caption         =   "�˻�"
         Height          =   375
         Left            =   4680
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtIndex 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Text            =   "�������"
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Label txtSuggest 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   840
      Width           =   6975
   End
   Begin VB.Label txtTotalPage 
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   9
      Top             =   8640
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   8
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label txtCurrentPage 
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   8640
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   ">  �˻�� �˻��ϰ� ���ϴ� �ּ����� ���� �����մϴ�."
      Height          =   255
      Left            =   6600
      TabIndex        =   6
      Top             =   480
      Width           =   5055
   End
End
Attribute VB_Name = "searchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public linkID As String
Public SecretKey As String
Public pageNum As Integer
Public suggestIndex As String

Private JusolinkService As New Jusolink

'�˻� ��ư
Private Sub btnSearch_Click()
    Dim searchInfo As SearchResult
    Dim perPage As Integer
    Dim noSuggest As Boolean
    Dim noDiffer As Boolean
    Dim relatedJibun As String
    Dim i As Integer
    Dim k As Integer
    
    txtIndex.SetFocus
    
    perPage = 20            '������ ��ϰ���
    noDiffer = False        '����˻� ����
    noSuggest = False        '�������þ� ����
    
    If pageNum = 0 Then
        pageNum = 1             '������ ��ȣ
    End If
    
    Set searchInfo = JusolinkService.Search(txtIndex.Text, pageNum, perPage, noDiffer, noSuggest)
    
    If searchInfo Is Nothing Then
        MsgBox ("[" + CStr(JusolinkService.LastErrCode) + "] " + JusolinkService.LastErrMessage)
        
        ListView1.ListItems.Clear
        txtCurrentPage.Caption = 0
        txtTotalPage.Caption = 0
        Exit Sub
    End If
    
    If searchInfo.numFound > 0 Then     '�� �˻���� ��
        
        pageNum = CInt(searchInfo.page)                 '�˻� ������ ��ȣ
        txtCurrentPage.Caption = searchInfo.page
        txtTotalPage.Caption = searchInfo.totalPage     '�� ������ ��ȣ
        
        '�������þ�
        If Len(searchInfo.suggest) > 0 Then
            txtSuggest.Caption = "�������þ� : " + searchInfo.suggest + " �˻���� ����"
            suggestIndex = searchInfo.suggest
        Else
            txtSuggest.Caption = ""
        End If
        
        '����Ʈ��
        ListView1.ListItems.Clear
                    
        If (searchInfo.juso Is Nothing) = False Then
            For i = 1 To searchInfo.juso.Count
                ListView1.ListItems.Add i, , searchInfo.juso.Item(i).zipcode        '�����ȣ
                ListView1.ListItems(i).SubItems(1) = searchInfo.juso.Item(i).sectionNum     '�������ȣ
                ListView1.ListItems(i).SubItems(2) = searchInfo.juso.Item(i).roadAddr1      '���θ��ּ�
                ListView1.ListItems(i).SubItems(3) = searchInfo.juso.Item(i).roadAddr2      '���θ��ּ� �ΰ�����
                ListView1.ListItems(i).SubItems(4) = searchInfo.juso.Item(i).jibunAddr      '�����ּ�
               
                
                If (searchInfo.juso.Item(i).relatedJibun Is Nothing) = False Then
                    For k = 1 To searchInfo.juso.Item(i).relatedJibun.Count
                        relatedJibun = relatedJibun + searchInfo.juso.Item(i).relatedJibun.Item(k) + " "
                    Next
                End If
                
                ListView1.ListItems(i).SubItems(5) = relatedJibun '��������
            Next
        
        End If
    Else
        txtSuggest.Caption = "�˻������ �����ϴ�."
    End If
    
    
End Sub
'ListView ������ Ŭ��
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    detailForm.zipcode = ListView1.SelectedItem.Text    '�����ȣ
    detailForm.sectionNum = ListView1.SelectedItem.SubItems(1) '�������ȣ
    detailForm.roadAddr1 = ListView1.SelectedItem.SubItems(2) ' ���θ��ּ�
    detailForm.roadAddr2 = ListView1.SelectedItem.SubItems(3) ' ���θ��ּ� �ΰ�����
    detailForm.jibunAddr = ListView1.SelectedItem.SubItems(4) ' �����ּ�
        
    detailForm.Show
    detailForm.txtRoadDetail.SetFocus
    Unload Me
End Sub
'����������
Private Sub btnPrevPage_Click()
    If pageNum > 1 Then
        pageNum = pageNum - 1
        Call btnSearch_Click
    End If
End Sub
'����������
Private Sub btnNextPage_Click()
    If pageNum < CInt(txtTotalPage.Caption) Then
        pageNum = pageNum + 1
        Call btnSearch_Click
    End If
End Sub
'�������þ� Ŭ��
Private Sub txtSuggest_Click()
    If Len(suggestIndex) > 0 Then
        txtIndex.Text = suggestIndex
        Call btnSearch_Click
    End If
End Sub
Private Sub Form_Load()
    '�ּҸ�ũ��� �ʱ�ȭ
    JusolinkService.Initialize linkID, SecretKey
    
    ListView1.FullRowSelect = True
        
    '����Ʈ�� Row Height
    ImageList1.ImageHeight = 21
    ImageList1.ImageWidth = 1
    ImageList1.ListImages.Add , , Me.Icon
    Set ListView1.SmallIcons = ImageList1
    
    If Len(txtIndex.Text) > 0 Then
        txtIndex.SelStart = Len(txtIndex.Text)
    End If
    
End Sub
Private Sub txtIndex_Change()
    pageNum = 1
End Sub

Private Sub txtIndex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call btnSearch_Click
    End If
End Sub
