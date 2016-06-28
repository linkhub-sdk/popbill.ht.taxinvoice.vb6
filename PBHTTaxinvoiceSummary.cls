VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBHTTaxinvoiceSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public count            As Long  '수집 결과 건수
Public supplyCostTotal  As Long  '공급가액 합계
Public taxTotal         As Long  '세액 합계
Public amountTotal      As Long  '합계 금액

Public Sub fromDictionary(dic As Variant)

    count = dic.Item("count")
    supplyCostTotal = dic.Item("supplyCostTotal")
    taxTotal = dic.Item("taxTotal")
    amountTotal = dic.Item("amountTotal")
    
End Sub
