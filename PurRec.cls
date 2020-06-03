VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PurRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is Industry Bill Interface Class, made by K3BOSPLUGINSWIZAED"
 
'定义插件对象接口. 必须具有的声明, 以此来获得事件
Private WithEvents m_BillTransfer   As K3BillTransfer.Bill
Attribute m_BillTransfer.VB_VarHelpID = -1
Private headDic As Scripting.Dictionary, entryDic As Scripting.Dictionary
 
Public Sub Show(ByVal oBillTransfer As Object)
 
    '接口实现
    '注意: 此方法必须存在, 请勿修改
    Set m_BillTransfer = oBillTransfer
    Set headDic = New Scripting.Dictionary
    Set entryDic = New Scripting.Dictionary
    Dim i As Long
    For i = LBound(m_BillTransfer.HeadCtl) To UBound(m_BillTransfer.HeadCtl)
        headDic.Add m_BillTransfer.HeadCtl(i).FieldName, i
    Next
    For i = LBound(m_BillTransfer.EntryCtl) To UBound(m_BillTransfer.EntryCtl)
        entryDic.Add m_BillTransfer.EntryCtl(i).FieldName, i
    Next
End Sub

Private Sub Class_Terminate()
 
    '释放接口对象
    '注意: 此方法必须存在, 请勿修改
    Set m_BillTransfer = Nothing

End Sub

Private Sub splitTransFee()
On Error GoTo Handle
    If Date > "2021-06-01" Then Exit Sub
    Dim fee As Double, tamount As Double, i As Long, price As Double, amount As Double, qty As Double
    fee = Val(m_BillTransfer.GetHeadText(headDic("FHeadSelfA0141")))
    If fee = 1 Then fee = 0
    i = 1
    While m_BillTransfer.GetGridText(i, entryDic("FItemID")) <> ""
        tamount = tamount + Val(m_BillTransfer.GetGridText(i, entryDic("Famount")))
        i = i + 1
    Wend
    If tamount = 0 Then Exit Sub
    i = 1
    While m_BillTransfer.GetGridText(i, entryDic("FItemID")) <> ""
        qty = Val(m_BillTransfer.GetGridText(i, entryDic("Fauxqty")))
        price = Val(m_BillTransfer.GetGridText(i, entryDic("Fauxprice")))
        amount = Val(m_BillTransfer.GetGridText(i, entryDic("Famount")))
        m_BillTransfer.SetGridText i, entryDic("FEntrySelfA0159"), price
        m_BillTransfer.SetGridText i, entryDic("FEntrySelfA0160"), fee * amount / tamount
        If qty <> 0 Then
            m_BillTransfer.SetGridText i, entryDic("FEntrySelfA0161"), fee * amount / tamount / qty + price
        Else
            m_BillTransfer.SetGridText i, entryDic("FEntrySelfA0161"), price
        End If
        i = i + 1
    Wend
    Exit Sub
Handle:
    MsgBox Err, vbOKOnly, "中基提示"
End Sub


Private Sub m_BillTransfer_GridChange(ByVal Col As Long, ByVal row As Long, ByVal value As Variant, ByVal bNewBill As Boolean, Cancel As Boolean)
    splitTransFee
End Sub

Private Sub m_BillTransfer_HeadChange(ByVal CtlIndex As Long, ByVal value As Variant, ByVal bNewBill As Boolean, Cancel As Boolean)
    If CtlIndex = headDic("FHeadSelfA0141") Then splitTransFee
End Sub
