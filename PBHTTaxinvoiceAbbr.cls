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

Public ntsconfirmNum    As String   '����û���ι�ȣ
Public writeDate        As String   '�ۼ�����
Public issueDate        As String   '��������
Public sendDate         As String   '��������
Public taxType          As String   '��������
Public purposeType      As String   '����/û��
Public supplyCostTotal  As String   '���ް��� �հ�
Public taxTotal         As String   '�����հ�
Public totalAmount      As String   '�հ�ݾ�
Public remark1          As String   '���

Public modifyYN         As Boolean  '���� ����(����)��꼭 ����
Public orgNTSConfirmNum   As String '���� ����(����)��꼭 ����û���ι�ȣ

Public purchaseDate     As String   '�ŷ�����
Public itemName         As String   'ǰ��
Public spec             As String   '�԰�
Public qty              As String   '����
Public unitCost         As String   '�ܰ�
Public supplyCost       As String   '���ް���
Public tax              As String   '����
Public remark           As String   '���

Public invoicerCorpNum  As String   '������ ����ڹ�ȣ
Public invoicerTaxRegID As String   '������ ��������ȣ
Public invoicerCorpName As String   '������ ��ȣ
Public invoicerCEOName  As String   '������ ��ǥ�ڼ���
Public invoicerEmail    As String   '������ ����� �̸���

Public invoiceeCorpNum  As String   '���޹޴��� ����ڹ�ȣ
Public invoiceeType     As String   '���޹޴��� ����
Public invoiceeTaxRegID As String   '���޹޴��� ��������ȣ
Public invoiceeCorpName As String   '���޹޴��� ��ȣ
Public invoiceeCEOName  As String   '���޹޴��� ��ǥ�� ����
Public invoiceeEmail1   As String   '���޹޴��� ����� �̸���
Public invoiceeEmail2   As String

Public trusteeCorpNum   As String   '��Ź�� ����ڹ�ȣ
Public trusteeTaxRegID  As String   '��Ź�� ��������ȣ
Public trusteeCorpName  As String   '��Ź�ڻ�ȣ
Public trusteeCEOName   As String   '��Ź�� ��ǥ�ڼ���
Public trusteeEmail     As String   '��Ź�� ����� �̸���

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
