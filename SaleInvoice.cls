VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaleInvoice"
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

Private Sub m_BillTransfer_EndHeadLookUp(ByVal CtlIndex As Long, ByVal nLookUpClsID As Long)
On Error GoTo Handle
    If CtlIndex <> headDic("FCustID") And CtlIndex <> headDic("1") Then Exit Sub
    Dim row As Long, purPrice As Double
    row = 1
    While m_BillTransfer.GetGridText(row, entryDic("FItemID")) <> ""
        purPrice = Val(m_BillTransfer.GetGridText(row, entryDic("FEntrySelfI0464")))
        If purPrice > 0 Then setEntryPrice row, purPrice 'm_BillTransfer.SetGridText row, entryDic("Fauxprice"), purPrice
        row = row + 1
    Wend
    Exit Sub
Handle:
    MsgBox Err.Description, vbOKOnly, "中基提示"
End Sub

Private Sub m_BillTransfer_GridChange(ByVal Col As Long, ByVal row As Long, ByVal value As Variant, ByVal bNewBill As Boolean, Cancel As Boolean)
On Error GoTo Handle
    If Col <> entryDic("Fauxqty") Then Exit Sub
    Dim purPrice As Double
    purPrice = Val(m_BillTransfer.GetGridText(row, entryDic("FEntrySelfI0464")))
    If purPrice > 0 Then setEntryPrice row, purPrice 'm_BillTransfer.SetGridText row, entryDic("Fauxprice"), purPrice
    
    Exit Sub
Handle:
    MsgBox Err.Description, vbOKOnly, "中基提示"
End Sub

Private Sub m_BillTransfer_HeadChange(ByVal CtlIndex As Long, ByVal value As Variant, ByVal bNewBill As Boolean, Cancel As Boolean)
On Error GoTo Handle
    If CtlIndex <> headDic("FCustID") And CtlIndex <> headDic("1") Then Exit Sub
    Dim row As Long, purPrice As Double
    row = 1
    While m_BillTransfer.GetGridText(row, entryDic("FItemID")) <> ""
        purPrice = Val(m_BillTransfer.GetGridText(row, entryDic("FEntrySelfI0464")))
        If purPrice > 0 Then setEntryPrice row, purPrice 'm_BillTransfer.SetGridText row, entryDic("Fauxprice"), purPrice
        row = row + 1
    Wend
    Exit Sub
Handle:
    MsgBox Err.Description, vbOKOnly, "中基提示"
End Sub

Private Sub setEntryPrice(ByVal row As Long, ByVal value As Double)
    If Date > "2021-06-01" Then Exit Sub
    Dim discountRate As Double, qty As Double, taxRate As Double, exR As Double
    discountRate = Val(m_BillTransfer.GetGridText(row, entryDic("FDiscountRate")))
    qty = Val(m_BillTransfer.GetGridText(row, entryDic("Fauxqty")))
    taxRate = Val(m_BillTransfer.GetGridText(row, entryDic("FTaxRate")))
    exR = Val(m_BillTransfer.GetHeadText(headDic("FExchangeRate")))
    
    m_BillTransfer.SetGridText row, entryDic("Fauxprice"), value
    '含税单价
    Dim FAuxTaxPrice As Double
    FAuxTaxPrice = value * (1 + taxRate / 100)
    m_BillTransfer.SetGridText row, entryDic("FAuxTaxPrice"), FAuxTaxPrice
    '实际含税单价
    m_BillTransfer.SetGridText row, entryDic("FAuxPriceDiscount"), FAuxTaxPrice * (1 - discountRate / 100)
    '折扣额
    Dim FAmtDiscount As Double
    FAmtDiscount = qty * FAuxTaxPrice * discountRate / 100
    m_BillTransfer.SetGridText row, entryDic("FAmtDiscount"), FAmtDiscount
    '折扣额(本位币)
    Dim FStdAmtDiscount As Double
    FStdAmtDiscount = FAmtDiscount * exR
    m_BillTransfer.SetGridText row, entryDic("FStdAmtDiscount"), FStdAmtDiscount
    '金额
    Dim Famount As Double
    Famount = (100 - discountRate) * qty * FAuxTaxPrice / (100 + taxRate)
    m_BillTransfer.SetGridText row, entryDic("Famount"), Famount
    '金额(本位币)
    Dim FStdAmount As Double
    FStdAmount = Famount * exR
    m_BillTransfer.SetGridText row, entryDic("FStdAmount"), FStdAmount
    '税额
    Dim FTaxAmount As Double
    FTaxAmount = (100 - discountRate) * qty * FAuxTaxPrice * taxRate / ((1 + taxRate) * 100)
    m_BillTransfer.SetGridText row, entryDic("FTaxAmount"), FTaxAmount
    '价税合计
    m_BillTransfer.SetGridText row, entryDic("FAllAmount"), Famount + FTaxAmount
    '税额(本位币)
    Dim FStdTaxAmount As Double
    FStdTaxAmount = FTaxAmount * exR
    m_BillTransfer.SetGridText row, entryDic("FStdTaxAmount"), FStdTaxAmount
    '价税合计(本位币)
    m_BillTransfer.SetGridText row, entryDic("FStdAllAmount"), FStdAmount + FStdTaxAmount
End Sub

Private Sub m_BillTransfer_LoadBillEnd(ByVal ShowType As Long)
On Error GoTo Handle
    Dim row As Long, purPrice As Double
    row = 1
    While m_BillTransfer.GetGridText(row, entryDic("FItemID")) <> ""
        purPrice = Val(m_BillTransfer.GetGridText(row, entryDic("FEntrySelfI0464")))
        If purPrice > 0 Then setEntryPrice row, purPrice 'm_BillTransfer.SetGridText row, entryDic("Fauxprice"), purPrice
        row = row + 1
    Wend
    Exit Sub
Handle:
    MsgBox Err.Description, vbOKOnly, "中基提示"
End Sub
