VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBHTTaxinvoiceXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ResultCode     As Long    '응답 상태코드
Public Message        As String  '전자(세금)계산서 국세청승인번호
Public retObject      As String  '전자(세금)계산서 XML 문서

Public Sub fromDictionary(dic As Variant)
    ResultCode = dic.Item("ResultCode")
    Message = dic.Item("Message")
    retObject = dic.Item("retObject")
End Sub

