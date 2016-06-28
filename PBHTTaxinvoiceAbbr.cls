VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBHTTaxinvoiceAbbr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ntsconfirmNum    As String   '국세청승인번호
Public writeDate        As String   '작성일자
Public issueDate        As String   '발행일자
Public sendDate         As String   '전송일자
Public taxType          As String   '과세형태
Public purposeType      As String   '영수/청구
Public supplyCostTotal  As String   '공급가액 합계
Public taxTotal         As String   '세액합계
Public totalAmount      As String   '합계금액
Public remark1          As String   '비고

Public modifyYN         As Boolean  '수정 전자(세금)계산서 여부
Public orgNTSConfirmNum   As String '원본 전자(세금)계산서 국세청승인번호

Public purchaseDate     As String   '거래일자
Public itemName         As String   '품명
Public spec             As String   '규격
Public qty              As String   '수량
Public unitCost         As String   '단가
Public supplyCost       As String   '공급가액
Public tax              As String   '세액
Public remark           As String   '비고

Public invoicerCorpNum  As String   '공급자 사업자번호
Public invoicerTaxRegID As String   '공급자 종사업장번호
Public invoicerCorpName As String   '공급자 상호
Public invoicerCEOName  As String   '공급자 대표자성명
Public invoicerEmail    As String   '공급자 담당자 이메일

Public invoiceeCorpNum  As String   '공급받는자 사업자번호
Public invoiceeType     As String   '공급받는자 유형
Public invoiceeTaxRegID As String   '공급받는자 종사업장번호
Public invoiceeCorpName As String   '공급받는자 상호
Public invoiceeCEOName  As String   '공급받는자 대표자 성명
Public invoiceeEmail1   As String   '공급받는자 담당자 이메일
Public invoiceeEmail2   As String

Public trusteeCorpNum   As String   '수탁자 사업자번호
Public trusteeTaxRegID  As String   '수탁자 종사업장번호
Public trusteeCorpName  As String   '수탁자상호
Public trusteeCEOName   As String   '수탁자 대표자성명
Public trusteeEmail     As String   '수탁자 담당자 이메일

Public Sub fromDictionary(dic As Variant)

    ntsconfirmNum = dic.Item("ntsconfirmNum")
    writeDate = dic.Item("writeDate")
    issueDate = dic.Item("issueDate")
    sendDate = dic.Item("sendDate")
    taxType = dic.Item("taxType")
    purposeType = dic.Item("purposeType")
    supplyCostTotal = dic.Item("supplyCostTotal")
    taxTotal = dic.Item("taxTotal")
    totalAmount = dic.Item("totalAmount")
    remark1 = dic.Item("remark1")
    
    modifyYN = dic.Item("modifyYN")
    orgNTSConfirmNum = dic.Item("orgNTSConfirmNum")
    
    purchaseDate = dic.Item("purchaseDate")
    itemName = dic.Item("itemName")
    spec = dic.Item("spec")
    qty = dic.Item("qty")
    unitCost = dic.Item("unitCost")
    supplyCost = dic.Item("supplyCost")
    tax = dic.Item("tax")
    remark = dic.Item("remark")
    
    invoicerCorpNum = dic.Item("invoicerCorpNum")
    invoicerTaxRegID = dic.Item("invoicerTaxRegID")
    invoicerCorpName = dic.Item("invoicerCorpName")
    invoicerCEOName = dic.Item("invoicerCEOName")
    invoicerEmail = dic.Item("invoicerEmail")
    
    
    invoiceeCorpNum = dic.Item("invoiceeCorpNum")
    invoiceeType = dic.Item("invoiceeType")
    invoiceeTaxRegID = dic.Item("invoiceeTaxRegID")
    invoiceeCorpName = dic.Item("invoiceeCorpName")
    invoiceeCEOName = dic.Item("invoiceeCEOName")
    invoiceeEmail1 = dic.Item("invoiceeEmail1")
    invoiceeEmail2 = dic.Item("invoiceeEmail2")
    
    trusteeCorpNum = dic.Item("trusteeCorpNum")
    trusteeTaxRegID = dic.Item("trusteeTaxRegID")
    trusteeCorpName = dic.Item("trusteeCorpName")
    trusteeCEOName = dic.Item("trusteeCEOName")
    trusteeEmail = dic.Item("trusteeEmail")
    
  
End Sub

