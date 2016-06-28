VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBHTTaxinvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public writeDate            As String
Public issueDT              As String
Public invoiceType          As String
Public taxType              As String
Public taxTotal             As String
Public supplyCostTotal      As String
Public totalAmount          As String
Public purposeType          As String
Public cash                 As String
Public chkBill              As String
Public credit               As String
Public note                 As String
Public remark1              As String
Public remark2              As String
Public remark3              As String
Public ntsconfirmNum        As String
Public modifyCode           As Integer
Public orgNTSConfirmNum     As String

Public invoicerCorpNum      As String
Public invoicerMgtKey       As String
Public invoicerTaxRegID     As String
Public invoicerCorpName     As String
Public invoicerCEOName      As String
Public invoicerAddr         As String
Public invoicerBizType      As String
Public invoicerBizClass     As String
Public invoicerContactName  As String
Public invoicerDeptName     As String
Public invoicerTEL          As String
Public invoicerEmail        As String

Public invoiceeCorpNum      As String
Public invoiceeType         As String
Public invoiceeMgtKey       As String
Public invoiceeTaxRegID     As String
Public invoiceeCorpName     As String
Public invoiceeCEOName       As String
Public invoiceeAddr         As String
Public invoiceeBizType      As String
Public invoiceeBizClass     As String
Public invoiceeContactName1 As String
Public invoiceeDeptName1    As String
Public invoiceeTEL1         As String
Public invoiceeEmail1       As String
Public invoiceeContactName2 As String
Public invoiceeDeptName2    As String
Public invoiceeTEL2         As String
Public invoiceeEmail2       As String

Public trusteeCorpNum       As String
Public trusteeMgtKey        As String
Public trusteeTaxRegID      As String
Public trusteeCorpName      As String
Public trusteeCEOName       As String
Public trusteeAddr          As String
Public trusteeBizType       As String
Public trusteeBizClass      As String
Public trusteeContactName   As String
Public trusteeDeptName      As String
Public trusteeTEL           As String
Public trusteeEmail         As String

Public detailList   As Collection

Public Sub fromDictionary(dic As Variant)

    writeDate = dic.Item("writeDate")
    issueDT = dic.Item("issueDT")
    invoiceType = dic.Item("invoiceType")
    taxType = dic.Item("taxType")
    taxTotal = dic.Item("taxTotal")
    supplyCostTotal = dic.Item("supplyCostTotal")
    totalAmount = dic.Item("totalAmount")
    purposeType = dic.Item("purposeType")
    cash = dic.Item("cash")
    chkBill = dic.Item("chkBill")
    credit = dic.Item("credit")
    note = dic.Item("note")
    remark1 = dic.Item("remark1")
    remark2 = dic.Item("remark2")
    remark3 = dic.Item("remark3")
    ntsconfirmNum = dic.Item("ntsconfirmNum")
    modifyCode = dic.Item("modifyCode")
    orgNTSConfirmNum = dic.Item("orgNTSConfirmNum")
    
    invoicerCorpNum = dic.Item("invoicerCorpNum")
    invoicerMgtKey = dic.Item("invoicerMgtKey")
    invoicerTaxRegID = dic.Item("invoicerTaxRegID")
    invoicerCorpName = dic.Item("invoicerCorpName")
    invoicerCEOName = dic.Item("invoicerCEOName")
    invoicerAddr = dic.Item("invoicerAddr")
    invoicerBizType = dic.Item("invoicerBizType")
    invoicerBizClass = dic.Item("invoicerBizClass")
    invoicerContactName = dic.Item("invoicerContactName")
    invoicerDeptName = dic.Item("invoicerDeptName")
    invoicerTEL = dic.Item("invoicerTEL")
    invoicerEmail = dic.Item("invoicerEmail")
    
    invoiceeCorpNum = dic.Item("invoiceeCorpNum")
    invoiceeType = dic.Item("invoiceeType")
    invoiceeMgtKey = dic.Item("invoiceeMgtKey")
    invoiceeTaxRegID = dic.Item("invoiceeTaxRegID")
    invoiceeCorpName = dic.Item("invoiceeCorpName")
    invoiceeCEOName = dic.Item("invoiceeCEOName")
    invoiceeAddr = dic.Item("invoiceeAddr")
    invoiceeBizType = dic.Item("invoiceeBizType")
    invoiceeBizClass = dic.Item("invoiceeBizClass")
    invoiceeContactName1 = dic.Item("invoiceeContactName1")
    invoiceeDeptName1 = dic.Item("invoiceeDeptName1")
    invoiceeTEL1 = dic.Item("invoiceeTEL1")
    invoiceeEmail1 = dic.Item("invoiceeEmail1")
    invoiceeContactName2 = dic.Item("invoiceeContactName2")
    invoiceeDeptName2 = dic.Item("invoiceeDeptName2")
    invoiceeTEL2 = dic.Item("invoiceeTEL2")
    invoiceeEmail2 = dic.Item("invoiceeEmail2")
            
    trusteeCorpNum = dic.Item("trusteeCorpNum")
    trusteeMgtKey = dic.Item("trusteeMgtKey")
    trusteeTaxRegID = dic.Item("trusteeTaxRegID")
    trusteeCorpName = dic.Item("trusteeCorpName")
    trusteeCEOName = dic.Item("trusteeCEOName")
    trusteeAddr = dic.Item("trusteeAddr")
    trusteeBizType = dic.Item("trusteeBizType")
    trusteeBizClass = dic.Item("trusteeBizClass")
    trusteeContactName = dic.Item("trusteeContactName")
    trusteeDeptName = dic.Item("trusteeDeptName")
    trusteeTEL = dic.Item("trusteeTEL")
    trusteeEmail = dic.Item("trusteeEmail")
    
    Dim tCollection As Collection
    
    Set tCollection = dic.Item("detailList")
    
    If (tCollection Is Nothing) = False Then
        Set detailList = New Collection
        
        Dim T As Variant
        Dim detail As PBHTTaxinvoiceDetail
        
        For Each T In tCollection
            Set detail = New PBHTTaxinvoiceDetail
            
            detail.fromDictionary T
            detailList.Add detail
            
        Next
    End If
    
End Sub


