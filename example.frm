VERSION 5.00
Begin VB.Form Example 
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16950
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   16950
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtPerPage 
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Text            =   "20"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtPageNum 
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Text            =   "1"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtIndex 
      Height          =   390
      Left            =   3480
      TabIndex        =   4
      Text            =   "용봉동"
      Top             =   360
      Width           =   3735
   End
   Begin VB.TextBox txtResult 
      Height          =   8655
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1920
      Width           =   13695
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "주소검색"
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton btnGetBalance 
      Caption         =   "잔여 포인트 조회"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton btnUnitCost 
      Caption         =   "검색 단가 조회"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "목록 갯수 :"
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "페이지 번호 :"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Example"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'링크아이디
Private Const linkID = "TESTER_JUSO"
'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "FjaRgAfVUPvSDHTrdd/uw/dt/Cdo3GgSFKyE1+NQ+bc="

Private JusolinkService As New Jusolink
'잔여포인트 확인
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = JusolinkService.GetBalance()
    
    If balance < 0 Then
        MsgBox ("[" + CStr(JusolinkService.LastErrCode) + "] " + JusolinkService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
End Sub
'주소검색
Private Sub btnSearch_Click()
    Dim searchInfo As SearchResult
    Dim i As Integer
    Dim k As Integer
    
    Set searchInfo = JusolinkService.Search(txtIndex.Text, CInt(txtPageNum.Text), CInt(txtPerPage.Text), False, True)
    
    If searchInfo Is Nothing Then
        MsgBox ("[" + CStr(JusolinkService.LastErrCode) + "] " + JusolinkService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "searches : " + searchInfo.searches + vbCrLf
    tmp = tmp + "numFound : " + searchInfo.numFound + vbCrLf
    tmp = tmp + "deletedWord : "
        
    If (searchInfo.deletedWord Is Nothing) = False Then
        For i = 1 To searchInfo.deletedWord.Count
            tmp = tmp + searchInfo.deletedWord.Item(i) + " "
        Next
    End If
    tmp = tmp + vbCrLf
    tmp = tmp + "suggest : " + searchInfo.suggest + vbCrLf
    tmp = tmp + "listSize : " + searchInfo.listSize + vbCrLf
    tmp = tmp + "totalPage : " + searchInfo.totalPage + vbCrLf
    tmp = tmp + "page : " + searchInfo.page + vbCrLf
    
    If (searchInfo.sidoCount Is Nothing) = False Then
        tmp = tmp + "GWANGJU : " + CStr(searchInfo.sidoCount.GWANGJU) + vbCrLf
    End If
    
    tmp = tmp + "chargeYN : " + CStr(searchInfo.chargeYN) + vbCrLf
        
    If (searchInfo.juso Is Nothing) = False Then
        For i = 1 To searchInfo.juso.Count
            tmp = tmp + "[ " + CStr(i) + " ] "
            tmp = tmp + searchInfo.juso.Item(i).zipcode + " "
            tmp = tmp + searchInfo.juso.Item(i).sectionNum + " "
            tmp = tmp + searchInfo.juso.Item(i).roadAddr1 + " "
            tmp = tmp + searchInfo.juso.Item(i).roadAddr2 + " "
            tmp = tmp + searchInfo.juso.Item(i).jibunAddr + " "
            'tmp = tmp + searchInfo.juso.Item(i).dongCode + " "
            'tmp = tmp + searchInfo.juso.Item(i).streetCode + " "
            
            If (searchInfo.juso.Item(i).detailBuildingName Is Nothing) = False Then
                For k = 1 To searchInfo.juso.Item(i).detailBuildingName.Count
                    tmp = tmp + searchInfo.juso.Item(i).detailBuildingName.Item(k) + " "
                Next
            End If
            
            If (searchInfo.juso.Item(i).relatedJibun Is Nothing) = False Then
                For k = 1 To searchInfo.juso.Item(i).relatedJibun.Count
                    tmp = tmp + searchInfo.juso.Item(i).relatedJibun.Item(k) + " "
                Next
            End If
            
            tmp = tmp + vbCrLf
        Next
    
    End If
    
    txtResult.Text = tmp
    
End Sub
'검색단가 확인
Private Sub btnUnitCost_Click()
    Dim unitCost As Single
    
    unitCost = JusolinkService.GetUnitCost()
    
    If unitCost < 0 Then
        MsgBox ("[" + CStr(JusolinkService.LastErrCode) + "] " + JusolinkService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "검색 단가 : " + CStr(unitCost)
    
End Sub

Private Sub Form_Load()
    JusolinkService.Initialize linkID, SecretKey
End Sub


