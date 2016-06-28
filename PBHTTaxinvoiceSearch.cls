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


Public code          As Long       '�����ڵ�
Public Message       As String     '���� �޽���
Public total         As Integer    '�� �˻���� �Ǽ�
Public perPage       As Integer    '�������� �˻�����
Public pageNum       As Integer    '������ ��ȣ
Public pageCount     As Integer    '������ ����

Public list As Collection          '����(����)��꼭 ���, PBHTTaxinvoiceAbbr.cls ����


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

