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

Public ResultCode     As Long    '���� �����ڵ�
Public Message        As String  '����(����)��꼭 ����û���ι�ȣ
Public retObject      As String  '����(����)��꼭 XML ����

Public Sub fromDictionary(dic As Variant)
    ResultCode = dic.Item("ResultCode")
    Message = dic.Item("Message")
    retObject = dic.Item("retObject")
End Sub
