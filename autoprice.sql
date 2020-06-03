/*
select * from t_ItemPropDesc where FItemClassID in (1,4) and FSQLColumnName like 'f_1%'
select * from t_ItemPropDesc where FItemClassID in (1,4) and FSearch in (501,10001)
select * from t_SubMessage where FTypeID=501


select * from sysobjects where type='u' and name like '%price%'

select * from t_TableDescription where FDescription like '%价格%'

select * from t_Price
select * from ICPrcPly
select * from ICPrcPlyEntry
select * from ICPrcPlyEntrySpec

select * from ICTransType
select * from ICTemplate where fid='i04'
select FFieldName,FHeadCaption,FRelationID,FAction--,* 
from ICTemplateEntry 
where FID='I04' and FValueType=1 and FLookUpCls=-1 and FRelationID like '%fauxprice%'


select FFieldName,FHeadCaption,FRelationID,FAction--,* 
from ICTemplateEntry 
where FID='I04' and FValueType=1 and FLookUpCls=-1 and FRelationID like '%fauxprice%'

select * from ICTemplateEntry where FID='a01'

--insert into t_ThirdPartyComponent
select * 
--delete i
from dbo.t_ThirdPartyComponent i 
where  fcomponentname like 'sl_zj_plus%' 
order by ftypedetailid,findex

delete t_ThirdPartyComponent where FTypeDetailID=80 and FIndex=1
insert into t_ThirdPartyComponent select 0,80,1,'sl_zj_plus.SaleInvoice','','中基取价'




--select * from t_Price
select * from ICPrcPly
select k.fnumber,j.fprice,* from ICPrcPlyEntry j join t_icitemcore k on k.fitemid=j.fitemid
select * from ICPrcPlyEntrySpec
go
*/
alter proc SL_AutoPrice --@purrecid int=0,@newpricetype int=0,@oldpricetype int=0,@itemid int=0
as
set nocount on
begin
	if GETDATE()>'2020-06-01' return
	declare @id int
	if not exists(select 1 from ICPrcPly where FPlyType='PrcAsm2')
	begin
		set @id=isnull((select MAX(FInterID) from ICPrcPly),0)+1
		insert into ICPrcPly(FNumber,FName,FPri,FPlyType,FSysTypeID,FPeriodType,FCycBegTime,FCycEndTime,FWeek,FMonth,FDayPerMonth,FSerialWeekPerMonth,FWeekDayPerMonth,FClassTypeID,FBillNo,FCycDesc,FBItem,FBCust,FBCustCls,FBEmp,FBVIPGrp,FBClientGrp,FBEmpCls,FClientGrp)
		select 'autoprice','自动价格',0,'PrcAsm2',501,0,'1900-01-01','1900-01-01','','',0,0,0,0,2,'',0,0,0,0,0,'',0,''
	end
	select @id=FInterID from ICPrcPly where FPlyType='PrcAsm2'
	create table #pur(FIdx int identity(1,1),FItemID int,FPrice decimal(28,10),FTransPrice decimal(28,10))
	insert into #pur(FItemID,FPrice,FTransPrice)
	select j.FItemID,j.FPrice,j.FEntrySelfA0160/j.FQty
	from ICStockBill i join ICStockBillEntry j on j.FInterID=i.FInterID
		join 
		(select j.FItemID,MAX(i.fdate) maxdate
		from ICStockBill i join ICStockBillEntry j on j.FInterID=i.FInterID
		where i.FTranType=1 /*and ISNULL(i.fcheckerid,0)<>0*/ and j.FAuxQty<>0
		group by j.FItemID) k on k.FItemID=j.FItemID and k.maxdate=i.FDate
	where i.FTranType=1 and j.FAuxQty<>0
	order by i.FDate,j.FEntryID
	
	delete i from #pur i left join (select FItemID,MAX(FIdx) maxid from #pur group by FItemID) j on j.FItemID=i.FItemID and j.maxid=i.FIdx where j.FItemID is null
	create table #ExPur(FRelatedID int,FItemID int,FPrice decimal(28,10))
	
	insert into #ExPur(FRelatedID,FItemID,FPrice)
	select k.F_101 FRelatedID,j.FItemID,round((i.FPrice+ISNULL(i.FTransPrice,0))*(1+isnull(k.F_103,0)/100),4) FPrice
	from #pur i join t_ICItemCustom j on j.FItemID=i.FItemID
		join t_Item_3004 k on k.F_102=j.F_102
	
	create table #price(FIdx int identity(1,1),FRelatedID int,FItemID int,FPrice decimal(28,10))
	insert into #price(FRelatedID,FItemID,FPrice)
	select i.FRelatedID,i.FItemID,i.FPrice from #ExPur i left join ICPrcPlyEntry j on j.FRelatedID=i.FRelatedID and j.FItemID=i.FItemID and j.FInterID=@id
	where i.FPrice<>isnull(j.FPrice,0)

	update i set FEndDate=DATEADD(DAY,-1,convert(char(10),GETDATE(),121))
	from ICPrcPlyEntry i join #price j on j.FRelatedID=i.FRelatedID and j.FItemID=i.FItemID
	
	insert into ICPrcPlyEntry(FInterID,FItemID,FRelatedID,FModel,FAuxPropID,FUnitID,FBegQty,FEndQty,FCuryID,FPriceType,FPrice,FBegDate,FEndDate,FLeadTime,FNote,FChecked,FIndex,FID,FBase,FBase1,FBegQty_Base,FEndQty_Base,FInteger,FClassTypeID,FBCust,FB2CustCls,FB2Emp,FB2EmpCls,FB2VipGrp,FB2Cust,FB2ItemID,FB2Item,FCreator,FOperator,FHelpCode)
	select @id,i.FItemID,i.FRelatedID,0,0,k.FUnitID,1,999999999,1,20004,i.FPrice,CONVERT(char(10),getdate(),121),'2100-01-01',0,'',1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,16393,16393,k1.FHelpCode
	from #price i join t_ICItembase k on k.FItemID=i.FItemID join t_ICItemCore k1 on k1.FItemID=i.FItemID
	
	insert into ICPrcPlyEntrySpec(FInterID,FItemID,FRelatedID,FLPriceCuryID,FLowPrice,FCanSell,FLPriceCtrl,FFlag)
	select @id,i.FItemID,i.FRelatedID,1,0,1,0,1
	from #price i left join ICPrcPlyEntrySpec j on j.FItemID=i.FItemID and j.FRelatedID=i.FRelatedID and j.FInterID=@id
	where j.FInterID is null
	drop table #pur
	drop table #ExPur
	drop table #price
end
go
SL_AutoPrice
go
alter trigger SL_TR_PurRec on icstockbill
for update,insert
as
set nocount on
--if exists(select 1 from inserted i join deleted j on j.FInterID=i.FInterID where ISNULL(i.FCheckDate,0)<>ISNULL(j.FCheckerID,0))
begin
	exec SL_AutoPrice
end
go
alter trigger SL_TR_Item on t_icitemcustom
for update,insert
as
set nocount on
if exists(select 1 from inserted i left join deleted j on j.FItemID=i.FItemID where ISNULL(i.F_102,0)<>ISNULL(j.F_102,0))
begin
	exec SL_AutoPrice
end
go
alter trigger SL_TR_PriceUp on t_item_3004
for insert,update
as
set nocount on
begin
	exec SL_AutoPrice
end
/*
select * from sysobjects where type='u' and name like '%max%'
select * from ICMaxNum where ftablename='ICPrcPly'
select * from t_stock
select * from ictemplateentry where fid='a01'
select * from t_itemclass
select * from t_itempropdesc where fitemclassid=4
select * from t_itempropdesc where fitemclassid=3004
*/