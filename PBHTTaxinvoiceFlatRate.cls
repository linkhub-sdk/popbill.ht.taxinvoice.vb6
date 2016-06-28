VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBHTTaxinvoiceFlatRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public referenceID      As String   '사업자번호
Public contractDT        As String   '정액제 서비스 시작일시
Public useEndDate         As String   '정액제 서비스 종료일
Public baseDate         As Integer  '자동연장 결제일
Public state            As Integer  '정액제 서비스 상태, 1-사용, 2-해지
Public closeRequestYN   As Boolean  '정액제 서비스 해지신청 여부
Public useRestrictYN    As Boolean  '정액제 서비스 사용제한 여부
Public closeOnExpired   As Boolean  '정액제 서비스 만료 시 해지 여부
Public unPaidYN         As Boolean  '미수금 보유 여부

Public Sub fromDictionary(dic As Variant)
    referenceID = dic.Item("referenceID")
    contractDT = dic.Item("contractDT")
    useEndDate = dic.Item("useEndDate")
    baseDate = dic.Item("baseDate")
    state = dic.Item("state")
    closeRequestYN = dic.Item("closeRequestYN")
    useRestrictYN = dic.Item("useRestrictYN")
    closeOnExpired = dic.Item("closeOnExpired")
    unPaidYN = dic.Item("unPaidYN")
End Sub



