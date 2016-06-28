VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBHTTaxinvoiceDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public serialNum    As Integer  '일련번호
Public purchaseDT   As String  '거래일자
Public itemName     As String   '품명
Public spec         As String   '규격
Public qty          As String   '수량
Public unitCost     As String   '단가
Public supplyCost   As String   '공급가액
Public tax          As String   '세액
Public remark       As String   '비고

Public Sub fromDictionary(dic As Variant)

    serialNum = dic.Item("serialNum")
    purchaseDT = dic.Item("purchaseDT")
    itemName = dic.Item("itemName")
    spec = dic.Item("spec")
    qty = dic.Item("qty")
    unitCost = dic.Item("unitCost")
    supplyCost = dic.Item("supplyCost")
    tax = dic.Item("tax")
    remark = dic.Item("remark")

End Sub



