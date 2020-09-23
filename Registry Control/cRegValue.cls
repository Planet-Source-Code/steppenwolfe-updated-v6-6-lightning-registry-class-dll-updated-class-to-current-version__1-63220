VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private m_Key        As String
Private m_DataType   As Reg_Type
Private m_Value      As Variant
Private m_sName      As Variant

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(ByVal PropVal As String)
    m_Key = PropVal
End Property

Public Property Get DataType() As Reg_Type
    DataType = m_DataType
End Property

Public Property Let DataType(PropVal As Reg_Type)
    m_DataType = PropVal
End Property

Public Property Get Value() As Variant
    Value = m_Value
End Property

Public Property Let Value(PropVal As Variant)
    m_Value = PropVal
End Property

Public Property Get sName() As Variant
    sName = m_sName
End Property

Public Property Let sName(PropVal As Variant)
    m_sName = PropVal
End Property

