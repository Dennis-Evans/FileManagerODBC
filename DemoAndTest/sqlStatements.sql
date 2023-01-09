-- create tables used by the tests, three small, simple tables
create table [dbo].[LabelDemo](
                  SysId int identity(1,1) not null, 
                  Label char(60) not null, 
                  Amount float not null constraint[DF_LabelDemo_Amount]  default(0), 
                  bitFlag bit null, 
                  guid uniqueidentifier null constraint [DF_LabelDemo_guid] default(newid()), 
                dateOf date null, timeof time(3) null 
               constraint [PK_LabelDemo] primary key clustered(SysId asc)); 
               create table [dbo].[Department](
                   SysId int identity(1,1) not null, 
                   LdSysId int not null, 
                   Label char(60) not null ); 
                 create table dbo.SecondLabelDemo(
                   SysId int identity(1,1) not null, 
                   Label char(60) not null,
                   Amount float not null  constraint [DF_Amount]  DEFAULT (0));
GO
-- add some test data 
insert into fmOdbctest.dbo.LabelDemo (Label, Amount, bitFlag)
              values('Fred', 12.33,  1), ('BARNEY', 14.12, 1), 
               ('Willma', 21.45, 1), ('Betty', 23.67, 0), 
               ('PEBBLES', 11.89, 0), ('Bam Bam', 2.34, 1);
              insert into fmOdbctest.dbo.Department(LdSysId, Label) 
               values(1, 'Big Rocks'), (2, 'Little Rocks'),  
                     (3, 'Big Rocks'), (4, 'Little Rocks'); 
             insert into fmOdbctest.dbo.SecondLabelDemo(Label, Amount) 
                values('Tom', 23.45), ('Dick', 12.32), ('Harry', 34.55); 
GO
-- create the table type used in the tests, create the type before the stored procedures
create type LabelDemoType as table (
      SysId int not null,
      Label char(60) not null,
      Amount float not null,
      RowAction int not null);
GO
-- create the scalar functions used by the test
create function dbo.getId 
(
@inLabel nchar(60)
) 

  returns  int 

as 

begin 

  declare @retv int; 
  
  select @retv = ld.SysId 
  from dbo.LabelDemo ld 
  where ld.Label = @inLabel; 
  
  if (@retv is null) begin 
    set @retv = 0; 
  end; 
  
  return @retv 

end;
GO
create procedure dbo.addLabelRow
(
@inLabel char(60), 
@inAmount float, 
@newId int out
)

as

begin 
  
  set nocount on;
  
  insert into dbo.LabelDemo(label, amount)  values(@inLabel, @inAmount); 
  select @newId = scope_identity(); 
  
  return;

end;
GO
create procedure dbo.CountDemoLabels
(
@outNum int out
) 

as

begin

  set nocount on;
   
  select @outNum = count(*) 
  from dbo.LabelDemo ld;
  
  return; 

end;
GO
create procedure dbo.InsertaTable
(
@inTable dbo.LabelDemoType readonly
)

as 

begin

  set nocount on;
  declare @rows int, @errMsg varchar(1000);
  merge dbo.labelDemo as target using
  (select i.sysId, i.label, i.amount, i.rowAction
   from @inTable i) as
   source(sysId, label, amount, rowAction) on
   target.sysId = source.sysId
   when matched and (source.RowAction = 3) then
     delete
   when matched and (source.rowAction = 2) then
     update set
      target.label = source.label, 
      target.amount = source.amount 
   when not matched and (source.rowAction = 1) then
      insert(label, amount)
      values(source.label, source.amount);

   return;

end;
GO
create procedure dbo.ReadLabelDemo 

as 

begin 
  set nocount on; 
  
  select ld.SysId, ld.Label, ld.amount from dbo.LabelDemo ld order by ld.Label desc; 
  
  return; 
end;
GO
create procedure dbo.ReadLabelDemoByLabel
(
@inLabel nchar(60)
) 

as 

begin 
  
  set nocount on; 
  
  select ld.sysId, rtrim(ld.Label), ld.Amount from dbo.LabelDemo ld where ld.Label = @inLabel order by ld.Label desc; 
 
  return; 

end;
GO
create procedure dbo.ReadLabelDemoWithCount
(
@outNumberRows int out
) 

as 

begin 

  set nocount on; 
  
  select ld.sysId, rtrim(ld.Label), ld.Amount from dbo.LabelDemo ld order by ld.Label; 
  select @outNumberRows = count(ld.SysId) from dbo.LabelDemo ld; 
  
  return; 
  
end;
GO
create procedure dbo.ReadOneRowLabel
(
@outSysId int out, 
@inLabel char(60), 
@outAmount float out
) 

as 

begin 

  set nocount on; 
  
   select @outSysid = ld.sysId, @outAmount = ld.Amount from dbo.LabelDemo ld where ld.Label = @inLabel; 
   
   return; 
   
end;
GO
create procedure dbo.ReadTwo
(
@inAmount float
) 

as  

begin 

  set nocount on; 
  
  select ld.SysId, ld.Label, ld.amount from dbo.LabelDemo ld; 
  
  select sld.Amount, sld.Label from dbo.SecondLabelDemo sld where sld.Amount > @inAmount; 
  
  select ld.Label, d.Label from dbo.Department d inner join dbo.LabelDemo ld on ld.SysId = d.LdSysId; 
  
  return; 

end;
GO