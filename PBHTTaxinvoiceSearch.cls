VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBHTTaxinvoiceSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public code          As Long       '응답코드
Public Message       As String     '응답 메시지
Public total         As Integer    '총 검색결과 건수
Public perPage       As Integer    '페이지당 검색개수
Public pageNum       As Integer    '페이지 번호
Public pageCount     As Integer    '페이지 개수

Public list As Collection          '전자(세금)계산서 목록, PBHTTaxinvoiceAbbr.cls 참고


Public Sub fromDictionary(dic As Variant)

    code = dic.Item("code")
    total = dic.Item("total")
    perPage = dic.Item("perPage")
    pageNum = dic.Item("pageNum")
    pageCount = dic.Item("pageCount")
    Message = dic.Item("message")
        
    Dim tCollection As Collection
    
    Set tCollection = dic.Item("list")
    
    If (tCollection Is Nothing) = False Then
        Set list = New Collection
        
        Dim T As Variant
        Dim tiList As PBHTTaxinvoiceAbbr

        For Each T In tCollection
            Set tiList = New PBHTTaxinvoiceAbbr
            
            tiList.fromDictionary T
            
            list.Add tiList
        Next
        
    End If
End Sub


