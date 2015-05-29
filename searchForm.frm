VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form searchForm 
   BorderStyle     =   1  '단일 고정
   Caption         =   " 주소검색 Example"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   17970
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton btnPrevPage 
      Caption         =   "이전 페이지"
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton btnNextPage 
      Caption         =   "다음 페이지"
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
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "우편번호"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "새우편번호"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "도로명주소"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "부가정보"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "지번주소"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "관련지번목록"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "주소검색"
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin VB.CommandButton btnSearch 
         Caption         =   "검색"
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
         Text            =   "테헤란로"
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Label txtSuggest 
      BeginProperty Font 
         Name            =   "굴림"
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
         Name            =   "굴림"
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
         Name            =   "굴림"
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
         Name            =   "굴림"
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
      Caption         =   ">  검색어를 검색하고 원하는 주소정보 셀을 선택합니다."
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

'검색 버튼
Private Sub btnSearch_Click()
    Dim searchInfo As SearchResult
    Dim perPage As Integer
    Dim noSuggest As Boolean
    Dim noDiffer As Boolean
    Dim relatedJibun As String
    Dim i As Integer
    Dim k As Integer
    
    txtIndex.SetFocus
    
    perPage = 20            '페이지 목록갯수
    noDiffer = False        '차등검색 끄기
    noSuggest = False        '수정제시어 끄기
    
    If pageNum = 0 Then
        pageNum = 1             '페이지 번호
    End If
    
    Set searchInfo = JusolinkService.Search(txtIndex.Text, pageNum, perPage, noDiffer, noSuggest)
    
    If searchInfo Is Nothing Then
        MsgBox ("[" + CStr(JusolinkService.LastErrCode) + "] " + JusolinkService.LastErrMessage)
        
        ListView1.ListItems.Clear
        txtCurrentPage.Caption = 0
        txtTotalPage.Caption = 0
        Exit Sub
    End If
    
    If searchInfo.numFound > 0 Then     '총 검색결과 수
        
        pageNum = CInt(searchInfo.page)                 '검색 페이지 번호
        txtCurrentPage.Caption = searchInfo.page
        txtTotalPage.Caption = searchInfo.totalPage     '총 페이지 번호
        
        '수정제시어
        If Len(searchInfo.suggest) > 0 Then
            txtSuggest.Caption = "수정제시어 : " + searchInfo.suggest + " 검색결과 보기"
            suggestIndex = searchInfo.suggest
        Else
            txtSuggest.Caption = ""
        End If
        
        '리스트뷰
        ListView1.ListItems.Clear
                    
        If (searchInfo.juso Is Nothing) = False Then
            For i = 1 To searchInfo.juso.Count
                ListView1.ListItems.Add i, , searchInfo.juso.Item(i).zipcode        '우편번호
                ListView1.ListItems(i).SubItems(1) = searchInfo.juso.Item(i).sectionNum     '새우편번호
                ListView1.ListItems(i).SubItems(2) = searchInfo.juso.Item(i).roadAddr1      '도로명주소
                ListView1.ListItems(i).SubItems(3) = searchInfo.juso.Item(i).roadAddr2      '도로명주소 부가정보
                ListView1.ListItems(i).SubItems(4) = searchInfo.juso.Item(i).jibunAddr      '지번주소
               
                
                If (searchInfo.juso.Item(i).relatedJibun Is Nothing) = False Then
                    For k = 1 To searchInfo.juso.Item(i).relatedJibun.Count
                        relatedJibun = relatedJibun + searchInfo.juso.Item(i).relatedJibun.Item(k) + " "
                    Next
                End If
                
                ListView1.ListItems(i).SubItems(5) = relatedJibun '관련지번
            Next
        
        End If
    Else
        txtSuggest.Caption = "검색결과가 없습니다."
    End If
    
    
End Sub
'ListView 아이템 클릭
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    detailForm.zipcode = ListView1.SelectedItem.Text    '우편번호
    detailForm.sectionNum = ListView1.SelectedItem.SubItems(1) '새우편번호
    detailForm.roadAddr1 = ListView1.SelectedItem.SubItems(2) ' 도로명주소
    detailForm.roadAddr2 = ListView1.SelectedItem.SubItems(3) ' 도로명주소 부가정보
    detailForm.jibunAddr = ListView1.SelectedItem.SubItems(4) ' 지번주소
        
    detailForm.Show
    detailForm.txtRoadDetail.SetFocus
    Unload Me
End Sub
'이전페이지
Private Sub btnPrevPage_Click()
    If pageNum > 1 Then
        pageNum = pageNum - 1
        Call btnSearch_Click
    End If
End Sub
'다음페이지
Private Sub btnNextPage_Click()
    If pageNum < CInt(txtTotalPage.Caption) Then
        pageNum = pageNum + 1
        Call btnSearch_Click
    End If
End Sub
'수정제시어 클릭
Private Sub txtSuggest_Click()
    If Len(suggestIndex) > 0 Then
        txtIndex.Text = suggestIndex
        Call btnSearch_Click
    End If
End Sub
Private Sub Form_Load()
    '주소링크모듈 초기화
    JusolinkService.Initialize linkID, SecretKey
    
    ListView1.FullRowSelect = True
        
    '리스트뷰 Row Height
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
