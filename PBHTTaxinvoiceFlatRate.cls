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

Public referenceID      As String   '����ڹ�ȣ
Public contractDT        As String   '������ ���� �����Ͻ�
Public useEndDate         As String   '������ ���� ������
Public baseDate         As Integer  '�ڵ����� ������
Public state            As Integer  '������ ���� ����, 1-���, 2-����
Public closeRequestYN   As Boolean  '������ ���� ������û ����
Public useRestrictYN    As Boolean  '������ ���� ������� ����
Public closeOnExpired   As Boolean  '������ ���� ���� �� ���� ����
Public unPaidYN         As Boolean  '�̼��� ���� ����

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


