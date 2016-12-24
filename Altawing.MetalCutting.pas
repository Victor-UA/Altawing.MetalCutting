Uses 'Victors', 'ProgressBar', 'iif';
Const
  CurrDate=Now;

Var
  SQL: String;
  i: Integer;
  n: Integer;
  k: Integer;
  SCount: Integer;
  SandDL: icmDictionaryList;
  GlOrdersDL: icmDictionaryList;
  ErrOrdersDL: icmDictionaryList;
  GlOrder: IdocGlassOrder;
  GlOrderNames: String='';
  GlOrderTitle: String;
  GlTask: IdocGlassTask;
  GlTaskKey: Variant=-100500;
  GlTaskCapacity: Integer=0;
  Customer: idocCustomer;
  Owner: idocEmployee;
  Err: String;
  PInfo: String;
  pInfoBegin: String;
  pInfoModify: String;
  pInfoEnd: String;
  tmp: String;
  PBegin, PEnd: Integer;
  ArtGlassD: icmDictionary;
  GrOrderKeys: String;
  GUID: String;

procedure OnCloseProgram;
begin
  try
    PBDestroy;
    GLOrdersDL.Clear;
    ErrOrdersDL.Clear;
    SandDL.Clear;
  except
//    Showmessage('Помилка під час завершення програми'+#13+ExceptionMessage);
  end;
end;

Begin
  PBCreate;
  Application.ProcessMessages;

  GLOrdersDL:=CreateDictionaryList;

    //-------------------------------------------------------------------------- Визначення списку id кроїв МПК
  i:=0;
  while i<Length(Documents) do begin
    GrOrderKeys:=GrOrderKeys+VarToStr(Documents[i].Key)+',';
    inc(i);
  end;
  GrOrderKeys:=LeftStr(GrOrderKeys,Length(GrOrderKeys)-1);

  //---------------------------------------------------------------------------- Перевірка замовлення на присутність в іншому крої
  SQL:='select'+#13+
       '  go1.name,'+#13+
       '  list (distinct o.orderno, '', '') Orders'+#13+
       'from grordersdetail gd'+#13+
       '  join orderitems oi on oi.orderitemsid=gd.orderitemsid'+#13+
       '  join orders o on o.orderid=oi.orderid'+#13+
       '  join order_uf_values ov on ov.orderid=o.orderid and ov.userfieldid=(select uf.userfieldid from userfields uf where uf.doctype=''IdocWindowOrder'' and uf.fieldname=''MetalSheetCutting'')'+#13+
       '  join grorders go1 on go1.guidhi=ov.var_guidhi and go1.guidlo = ov.var_guidlo'+#13+
       'where gd.grorderid in (:grorderid)'+#13+
       'group by 1';
  SQL:=ReplaceText(SQL, ':grorderid', GrOrderKeys);
  //showmessage(GrOrderKeys);//!
  try
    ErrOrdersDL:=QueryRecordList(SQL, MakeDictionary([]));
  except
    showmessage('Помилка отримання списку помилкових замовлень'+#13+ExceptionMessage+#13+SQL);
    OnCloseProgram;
    exit;
  end;
  try
    if ErrOrdersDL.Count>0 then begin
      Err:='Помилка створення крою: замовлення вже містять інші крої.'+#13;
      i:=0;
      while i<ErrOrdersDL.Count do begin
        Err:=Err+'Крой: ['+ErrOrdersDL[i]['name']+'], замовлення: ['+ErrOrdersDL[i]['Orders']+']'+#13;
        inc(i);
      end;
      Showmessage(Err);
      OnCloseProgram;
      exit;
    end;
  except
    showmessage('Системна помилка [замовлення вже містять інші крої]'+#13+ExceptionMessage);
    OnCloseProgram;
    exit;
  end;

  //---------------------------------------------------------------------------- Завантаження деталей з листового металу, ініціалізація id структури WindowOrder
  SQL:='select'+#13+
       '  d.goname,'+#13+
       '  d.orderno,'+#13+
       '  d.oiname,'+#13+
       '  d.marking,'+#13+
       '  d.name,'+#13+
       '  d.gggrgoodsid,'+#13+
       '  gg1.marking rmmarking,'+#13+
       '  gg1.name rmname,'+#13+
       '  gg1.grgoodsid grgoodsid,'+#13+
       '  ('+#13+
       '    select first 1'+#13+
       '      g.goodsid'+#13+
       '    from goods g'+#13+
       '    where g.grgoodsid=gg1.grgoodsid'+#13+
       '  ) goodsid,'+#13+
       '  ('+#13+
       '    select first 1'+#13+
       '      g.price1'+#13+
       '    from goods g'+#13+
       '    where g.grgoodsid=gg1.grgoodsid'+#13+
       '  ) price1,'+#13+
       ''+#13+
       '  d.width,'+#13+
       '  d.length,'+#13+
       '  d.qty,'+#13+
       '  d.rcomment,'+#13+
       '  d.itemsdetailid,'+#13+
       '  d.OrderExists,'+#13+
       '  ('+#13+
       '    select'+#13+
       '      gen_id(GEN_ITEMSDETAIL, 1)'+#13+
       '      from rdb$database'+#13+
       '  ) NewItemsDetailId,'+#13+
       '  d.grorderid,'+#13+
       '  0 as GlOrderId,'+#13+
       '  0 as OrderitemsId,'+#13+
       '  '''' as thumbs,'+#13+
       '  -1 as GlOrdersDLIndex'+#13+
       'from ('+#13+
       '  select'+#13+
       '    go.name goname,'+#13+
       '    o.orderno,'+#13+
       '    oi.name as oiname,'+#13+
       '    gg.marking,'+#13+
       '    gg.name,'+#13+
       '    gg.grgoodsid gggrgoodsid,'+#13+
       '    ('+#13+
       '      select first 1'+#13+
       '        gg1.grgoodsid'+#13+
       '      from groupgoods gg1'+#13+
       '        join groupgoodstypes ggt1 on ggt1.ggtypeid = gg1.ggtypeid and ggt1.code = ''MetalSheetLinear'''+#13+
       '      where gg1.width >= itd.width'+#13+
       '        and gg1.marking containing iif ('+#13+
       '          gg.marking containing '' ж/оц'', '' Оц '','+#13+
       '          iif('+#13+
       '            gg.marking containing '' б/п'', '' БП '','+#13+
       '            iif('+#13+
       '              gg.marking containing '' к/п'', '' КП '','+#13+
       '              '''''+#13+
       '            )'+#13+
       '          )'+#13+
       '        )'+#13+
       '        and'+#13+
       '        ('+#13+
       '          (itd.height<=1250 and itd.width<=1250)'+#13+
       '            or'+#13+
       '          (itd.height<=600 and itd.width>=1250 and itd.width<=2800)'+#13+
       '        )'+#13+
       ''+#13+
       '      order by gg1.width'+#13+
       '    ) rmgrgoodsid,'+#13+
       '    coalesce(itd.width,  0) length,'+#13+
       '    coalesce(itd.height,  0) width,'+#13+
       '    itd.qty qty,'+#13+
       '    coalesce(itd.rcomment, '''') rcomment,'+#13+
       '    itd.itemsdetailid,'+#13+
       '    iif(exists'+#13+
       '      (select *'+#13+
       '      from orders o'+#13+
       '      where o.orderno=''LinearMetal '' || go.name'+#13+
       '      ),1,0'+#13+
       '    ) OrderExists,'+#13+
       '  /*'+#13+
       '    ('+#13+
       '      select'+#13+
       '        gen_id(GEN_GRORDERSDETAIL, 1)'+#13+
       '        from rdb$database'+#13+
       '    ) grorderdetailid,'+#13+
       '  */'+#13+
       '    go.grorderid'+#13+
       '  from grordersdetail god'+#13+
       '    join grorders go on go.grorderid=god.grorderid'+#13+
       '    join orderitems oi on oi.orderitemsid=god.orderitemsid'+#13+
       '    join orders o on o.orderid=oi.orderid'+#13+
       '    join itemsdetail itd on itd.orderitemsid=oi.orderitemsid'+#13+
       '    join groupgoods gg on gg.grgoodsid=itd.grgoodsid'+#13+
       '    join groupgoodstypes ggt on ggt.ggtypeid = gg.ggtypeid and ggt.code = ''MetalSheetWare'''+#13+
       '  where go.grorderid in (:grorderid)'+#13+
       '    and god.isaddition=1'+#13+
       '  order by 1,2,3,4,5'+#13+
       ') d'+#13+
       '  join groupgoods gg1 on gg1.grgoodsid = d.rmgrgoodsid';
  SQL:=ReplaceText(SQL, ':grorderid', GrOrderKeys);
  try
    //showmessage(SQL);//!
    SandDL:=QueryRecordList(SQL, MakeDictionary([]));
    fPB.Max:=SandDL.Count;
    OldP:=0;
  except
    fPB.Max:=0;
    showmessage('Помилка запиту металу'+#13+ExceptionMessage+#13+SQL);
    OnCloseProgram;
    exit;
  end;

  //---------------------------------------------------------------------------- Ініціалізація id структури замовлень Металу
  if SandDL.Count>0 then begin
    SQL:='select'+#13+
         '  go.grorderid,'+#13+
         '  go.name,'+#13+
         '  go.groupdate,'+#13+
         '  ('+#13+
         '    select'+#13+
         '      gen_id(GEN_GRORDERSDETAIL, 1)'+#13+
         '      from rdb$database'+#13+
         '  ) grorderdetailid,'+#13+
         '  ( select'+#13+
         '      gen_id(GEN_ORDERS,1)'+#13+
         '    from rdb$database'+#13+
         '  ) as GlOrderId,'+#13+
         '  ( select'+#13+
         '      gen_id(GEN_APPROVEDOCUMENTS,1)'+#13+
         '    from rdb$database'+#13+
         '  ) as approvedocumentid,'+#13+
         '  ('+#13+
         '    select'+#13+
         '      gen_id(GEN_ORDERITEMS, 1)'+#13+
         '      from rdb$database'+#13+
         '  ) orderitemsid,'+#13+
         '  ( select'+#13+
         '      cu.customerid'+#13+
         '    from customers cu'+#13+
         '      join contragents ca on ca.contragid=cu.contragid'+#13+
         '    where ca.name=''Газда'''+#13+
         '  ) as customerid,'+#13+
         '  0 as isCreated'+#13+
         'from grorders go'+#13+
         'where go.grorderid in ('+GrOrderKeys+')';
    try
      GlOrdersDL:=QueryRecordList(SQL, MakeDictionary([]));
    except
      showmessage('Помилка визначення GlOrderId'+#13+ExceptionMessage+#13+SQL);
      OnCloseProgram;
      exit;
    end;
  end;

  n:=0;
  while n<SandDL.Count do begin
    i:=0;
    while (i<GlOrdersDL.Count) and (GlOrdersDL[i]['grorderid']<>SandDL[n]['grorderid']) do begin
      inc(i);
    end;
    if GlOrdersDL[i]['grorderid']<>SandDL[n]['grorderid'] then begin
      RaiseException('Не знайдено відповідного крою:'+#13+'GlOrdersDL['+VarToStr(i)+'].Key='+VarToStr(GlOrdersDL[i]['grorderid'])+'<>'+'SandDL['+VarToStr(n)+'][''grorderid'']');
    end;
    SandDL[n]['GlOrderId']:=GlOrdersDL[i]['GlOrderId'];
    SandDL[n]['orderitemsid']:=GlOrdersDL[i]['orderitemsid'];
    SandDL[n]['GlOrdersDLIndex']:=i;

    if GlTaskKey=-100500 then begin
      SQL:='select'+#13+
           '  gen_id(GEN_GRORDERS, 1)'+#13+
           'from rdb$database';
      try
        GlTaskKey:=QueryValue(SQL, MakeDictionary([]));
      except
        showmessage('Помилка визначення GrOrderId'+#13+ExceptionMessage+#13+SQL);
        OnCloseProgram;
        exit;
      end;

      //------------------------------------------------------------------------ Створення крою Металу
      SQL:='insert into grorders (grorderid, name, isoptimized, isdefault, makebill, rcomment, reccolor, recflag, guidhi, guidlo, ownerid, datecreated, datemodified, datedeleted, ownertype, procschemaid, groupdate, capacity, productionsid, planid, isclosed, deleted, linearoptim, linearopttype, linearuserest, linearsaverest, linearrestmode, linearpaired, layoutoptim, layoutopttype, layoutuserest, layoutsaverest, whlistid, ordernames)'+#13+
           'values ('+#13+
           '  :grorderid,'+#13+
           '  :name,'+#13+
           '  0,'+#13+
           '  0,'+#13+
           '  1,'+#13+
           '  :rcomment,'+#13+
           '  null,'+#13+
           '  null,'+#13+
           '  :guidhi,'+#13+
           '  :guidlo,'+#13+
           '  :ownerid,'+#13+
           '  :datecreated,'+#13+
           '  :datemodified,'+#13+
           '  null,'+#13+
           '  0,'+#13+
           '  null,'+#13+
           '  :groupdate,'+#13+
           '  :capacity,'+#13+       //Заповнити після виконання циклу
           '  null,'+#13+
           '  null,'+#13+
           '  0,'+#13+
           '  0,'+#13+
           '  1,'+#13+
           '  1,'+#13+
           '  1,'+#13+
           '  1,'+#13+
           '  null,'+#13+
           '  0,'+#13+
           '  0,'+#13+
           '  0,'+#13+
           '  0,'+#13+
           '  0,'+#13+
           '  0,'+#13+
           '  :ordernames'+#13+      //Заповнити після виконання циклу
           ')';
      GUID:=GenerateGUID;
      SQL:=ReplaceText(SQL, ':grorderid', VarToSQL(GlTaskKey));
      SQL:=ReplaceText(SQL, ':name', VarToSQL('Лінійний Листовий Метал від '+DateTimeToStr(CurrDate)));
      SQL:=ReplaceText(SQL, ':rcomment', VarToSQL('Створено автоматично'));
      SQL:=ReplaceText(SQL, ':guidhi', VarToSQL(GUIDHi(GUID)));
      SQL:=ReplaceText(SQL, ':guidlo', VarToSQL(GUIDLo(GUID)));
      SQL:=ReplaceText(SQl, ':ownerid', VarToSQL(UserContext.UserID));
      SQL:=ReplaceText(SQl, ':datecreated', VarToSQL(Now));
      SQL:=ReplaceText(SQl, ':datemodified', VarToSQL(Now));
      SQL:=ReplaceText(SQL, ':groupdate', VarToSQL(Now));
      SQL:=ReplaceText(SQL, ':capacity', VarToSQL(0));
      SQL:=ReplaceText(SQL, ':ordernames', VarToSQL(''));
      try
        Err:=ExecuteSQLCommit(SQL);
      except
        showmessage('Помилка створення змінного завдання склопакетів'+#13+ExceptionMessage+#13+SQL);
        OnCloseProgram;
        exit;
      end;
    end;

    if GlOrdersDL[i]['isCreated']=0 then begin
      try
        GlOrderTitle:='LinearMetal ' + GlOrdersDL[i]['name'];
        if SandDL[n]['OrderExists']>0 then begin
          GlOrderTitle:=GlOrderTitle+' ('+DateTimeToStr(CurrDate)+')';
        end;

        SQL:='insert into approvedocuments (approvedocumentid, doctype)'+#13+
             'values ('+#13+
             '  :approvedocumentid,'+#13+
             '  ''IdocOrder'''+#13+
             ')';
        SQL:=ReplaceText(SQL, ':approvedocumentid', VarToSQL(GlOrdersDL[i]['approvedocumentid']));
        Err:=ExecuteSQLCommit(SQL);
        if Err<>'' then begin
          showmessage('Помилка створення [approvedocuments]'+#13+Err);
          OnCloseProgram;
          exit;
        end;

        //---------------------------------------------------------------------- Створення замовлення Металу
        SQL:='insert into orders (orderid, ownertype, orderno, agreementno, agreementdate, currencyid, sellerid, customerid, itemstatusmode, totalpricelock, proddate, dateorder, orderstatus, lastgenitem, guidhi, guidlo, ownerid, datecreated, datemodified, deleted, rcomment, valid, totalprice, payment, isdealeradd, isdealerstartadd, isreserved, approvedocumentid, crossrate)'+#13+
             'values ('+#13+
             '  :orderid,'+#13+
             '  0,'+#13+
             '  :orderno,'+#13+
             '  null,'+#13+
             '  null,'+#13+
             '  ('+#13+
             '    select first 1'+#13+
             '      c.currencyid'+#13+
             '    from currency c'+#13+
             '    where c.isbase=1'+#13+
             '      and c.islogrecord=0'+#13+
             '      and c.deleted=0'+#13+
             '  ),'+#13+
             '  :sellerid,'+#13+
             '  :customerid,'+#13+
             '  0,'+#13+
             '  0,'+#13+
             '  :proddate,'+#13+
             '  :dateorder,'+#13+
             '  0,'+#13+
             '  0,'+#13+
             '  :guidhi,'+#13+
             '  :guidlo,'+#13+
             '  :ownerid,'+#13+
             '  :datecreated,'+#13+
             '  :datemodified,'+#13+
             '  0,'+#13+
             '  :rcomment,'+#13+
             '  1,'+#13+
             '  0,'+#13+
             '  0,'+#13+
             '  0,'+#13+
             '  0,'+#13+
             '  0,'+#13+
             '  :approvedocumentid,'+#13+
             '  1'+#13+
             ')';
        GUID:=GenerateGUID;
        SQL:=ReplaceText(SQl, ':orderid', VarToSQL(GlOrdersDL[i]['GlOrderId']));
        SQL:=ReplaceText(SQl, ':orderno', VarToSQL(GlOrderTitle));
        SQL:=ReplaceText(SQl, ':sellerid', VarToSQL('null'));
        SQL:=ReplaceText(SQl, ':customerid', VarToSQL(GlOrdersDL[i]['customerid']));
        SQL:=ReplaceText(SQl, ':proddate', VarToSQL(GlOrdersDL[i]['groupdate']));
        SQL:=ReplaceText(SQl, ':dateorder', VarToSQL(GlOrdersDL[i]['groupdate']));
        SQL:=ReplaceText(SQl, ':guidhi', VarToSQL(GUIDHi(GUID)));
        SQL:=ReplaceText(SQl, ':guidlo', VarToSQL(GUIDLo(GUID)));
        SQL:=ReplaceText(SQl, ':ownerid', VarToSQL(UserContext.UserID));
        SQL:=ReplaceText(SQl, ':datecreated', VarToSQL(Now));
        SQL:=ReplaceText(SQl, ':datemodified', VarToSQL(Now));
        SQL:=ReplaceText(SQl, ':rcomment', VarToSQL('Вибірка металу з крою ['+GlOrdersDL[i]['name']+']'));
        SQL:=ReplaceText(SQl, ':approvedocumentid', VarToSQL(GlOrdersDL[i]['approvedocumentid']));
        Err:=ExecuteSQLCommit(SQL);
        if Err<>'' then begin
          showmessage('Помилка створення замовлення склопакетів'+#13+Err);
          OnCloseProgram;
          exit;
        end;
        GlOrdersDL[i]['isCreated']:=1;
      except
        showmessage('Помилка створення замовлення склопакетів'+#13+ExceptionMessage+#13+GlOrderTitle);
        OnCloseProgram;
        exit;
      end;

        //-------------------------------------------------------------------------- Створення конструкції goods
      SQL:='insert into ORDERITEMS (orderitemsid, orderid, name, qty, laboriousness, area, rcomment, isaddition, usedqty, usedaddqty, thumbs, valid, price, costall, packinfo, productcount)'+#13+
           'values ('+#13+
           '  :orderitemsid,'+#13+
           '  :orderid,'+#13+
           '  ''goods'','+#13+
           '  1,'+#13+
           '  null,'+#13+
           '  null,'+#13+
           '  null,'+#13+
           '  1,'+#13+
           '  0,'+#13+
           '  0,'+#13+
           '  null,'+#13+
           '  0,'+#13+
           '  null,'+#13+
           '  0,'+#13+
           '  null,'+#13+
           '  0'+#13+
           ')';
      SQL:=ReplaceText(SQL, ':orderitemsid', VarToSQL(GlOrdersDL[i]['OrderItemsId']));
      SQL:=ReplaceText(SQL, ':orderid', VarToSQL(GlOrdersDL[i]['GlOrderId']));
      Err:=ExecuteSQLCommit(SQL);
      if Err<>'' then begin
        showmessage('Помилка створення конструкції'+#13+Err);
        OnCloseProgram;
        exit;
      end;

      //-------------------------------------------------------------------------- Створення деталізації крою Металу
      SQL:='insert into grordersdetail (grorderdetailid, grorderid, orderitemsid, qty, isaddition)'+#13+
           'values ('+#13+
           '  :grorderdetailid,'+#13+
           '  :grorderid,'+#13+
           '  :orderitemsid,'+#13+
           '  :qty,'+#13+
           '  :isaddition'+#13+
           ')';
      SQL:=ReplaceText(SQL, ':grorderdetailid', VarToSQL(GlOrdersDL[i]['grorderdetailid']));
      SQL:=ReplaceText(SQL, ':grorderid', VarToSQL(GlTaskKey));
      SQL:=ReplaceText(SQL, ':orderitemsid', VarToSQL(SandDL[n]['orderitemsid']));
//      SQL:=ReplaceText(SQL, ':qty', VarToSQL(SandDL[n]['qty']));
      SQL:=ReplaceText(SQL, ':qty', VarToSQL(1));
      SQL:=ReplaceText(SQL, ':isaddition', VarToSQL(1));
      Err:=ExecuteSQLCommit(SQL);
      if Err<>'' then begin
        showmessage('Помилка створення [grordersdetail]'+#13+Err);
        OnCloseProgram;
        exit;
      end;

    end;

    //-------------------------------------------------------------------------- Створення деталізації конструкції Металу
    SQL:='insert into itemsdetail (itemsdetailid,orderitemsid,grgoodsid,goodsid,modelno,width,height,thick,qty,ang1,ang2,radius,pricetype,weight,connection1,connection2,isextended,updatestatus,rcomment,allvolume,allsavingvolume,allweight,price,savingabs,cost,savingcost)'+#13+
         'values ('+#13+
         '  :itemsdetailid,'+#13+
         '  :orderitemsid,'+#13+
         '  :grgoodsid,'+#13+
         '  :goodsid,'+#13+
         '  0,'+#13+
         '  :width,'+#13+
         '  1,'+#13+
         '  :thick,'+#13+
         '  :qty,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  1,'+#13+
         '  0,'+#13+
         '  :rcomment,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  :price,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0'+#13+
         ')';
    SQL:=ReplaceText(SQL, ':itemsdetailid', VarToSQL(SandDL[n]['NewItemsDetailId']));
    SQL:=ReplaceText(SQL, ':orderitemsid', VarToSQL(SandDL[n]['orderitemsid']));
    SQL:=ReplaceText(SQL, ':grgoodsid', VarToSQL(SandDL[n]['grgoodsid']));
    SQL:=ReplaceText(SQL, ':goodsid', VarToSQL(SandDL[n]['goodsid']));
    SQL:=ReplaceText(SQL, ':width', VarToSQL(SandDL[n]['length']));
    SQL:=ReplaceText(SQL, ':thick', VarToSQL(SandDL[n]['width']));
    SQL:=ReplaceText(SQL, ':qty', VarToSQL(SandDL[n]['qty']));

    SQL:=ReplaceText(SQL, ':rcomment', VarToSQL(
      SandDL[n]['rcomment'] + #13 +
      SandDL[n]['marking'] + ' ' +
      VarToStr(
        iif(
          (SandDL[n]['rcomment'] = ''),
          VarToStr(
            iif(
              Pos('Отлив', SandDL[n]['name']) > 0 ,
              SandDL[n]['width'] - 50,
              iif(
                (Pos('Нестандарт', SandDL[n]['name']) = 0),
                SandDL[n]['width'] - 30,
                SandDL[n]['width']
              )
            )
          ) + 'x',
          //SandDL[n]['width']
          'L'
        )
      ) +
      VarToStr(SandDL[n]['length']) + #13 +
//      SandDL[n]['marking']+' '+SandDL[n]['name'] + #13 +
      '['+SandDL[n]['orderno']+'/'+SandDL[n]['oiname'] + ']'
    ));

//      SQL:=ReplaceText(SQL, ':name', VarToSQL(PADL(VarToStr(n),2,'0')+'['+copy(SandDL[n]['orderno']+'-'+SandDL[n]['oiname'],1,28)+']'));
    SQL:=ReplaceText(SQL, ':price', VarToSQL(SandDL[n]['price1']));
    Err:=ExecuteSQLCommit(SQL);
    if Err<>'' then begin
      showmessage('Помилка створення [itemsdetail]'+#13+Err);
      OnCloseProgram;
      exit;
    end;

    PBStep(SandDL[n]['marking']+'/'+SandDL[n]['goname']+'/'+SandDL[n]['orderno']+'/'+SandDL[n]['oiname']);

    GlTaskCapacity:=GlTaskCapacity+GlOrdersDL[i]['qty'];
    inc(n);
  end;

  FormPB.Hide;

  //---------------------------------------------------------------------------- Розрахунок крою Металу
  i:=0;
  while i<GlOrdersDL.Count do begin
    if GlOrdersDL[i]['isCreated']=1 then begin
      try
        GlOrder:=OpenDocument(IdocGlassOrder, GlOrdersDL[i]['GlOrderId']);
        GlOrder.Calculate;
        GlOrder.Save;

      except
        showmessage('Помилка калькуляції замовлення склопакетів ['+GlOrdersDL[i]['name']+']'+#13+ExceptionMessage);
        OnCloseProgram;
        exit;
      end;
    end;
    inc(i);
  end;
  GlOrder:=null;

  if GlTaskKey<>-100500 then begin
{
    while i<GlOrdersDL.Count do begin
      if GlOrdersDL[i]['isCreated']=1 then begin
        GlOrderNames:=GlOrderNames+GlOrdersDL[i]['name']+', ';
      end;
      inc(i);
    end;
    GlOrderNames:=Copy(GlOrderNames, 1, Length(GlOrderNames)-2);
}

    //-------------------------------------------------------------------------- Прив'язка крою Металу у замовленні конструкції
    SQL:='execute block'+#13+
         'as'+#13+
         '  declare variable orderid id;'+#13+
         'Begin'+#13+
         '  for'+#13+
         '    select'+#13+
         '      o.orderid'+#13+
         '    from orders o'+#13+
         '    where o.ownertype=0'+#13+
         '      and exists ('+#13+
         '        select'+#13+
         '          oi.orderid'+#13+
         '        from grordersdetail gd'+#13+
         '          join orderitems oi on oi.orderitemsid=gd.orderitemsid'+#13+
         '        where gd.grorderid in (:grorders)'+#13+
         '          and oi.orderid=o.orderid'+#13+
         '      )'+#13+
         '  into :orderid'+#13+
         '  do begin'+#13+
         '    update or insert into order_uf_values (orderid, userfieldid, var_str, var_guidhi, var_guidlo)'+#13+
         '    values ('+#13+
         '      :orderid,'+#13+
         '      (select uf.userfieldid from userfields uf where uf.fieldname=''MetalSheetCutting'' and uf.doctype=''IdocWindowOrder''),'+#13+
         '      (select go.name from grorders go where go.grorderid = :grorderid),'+#13+
         '      (select go.guidhi from grorders go where go.grorderid = :grorderid),'+#13+
         '      (select go.guidlo from grorders go where go.grorderid = :grorderid)'+#13+
         '    )'+#13+
         '    matching(orderid, userfieldid);'+#13+
         '  end'+#13+
         'End';
    SQL:=ReplaceText(SQL, ':grorders', GrOrderKeys);
    SQL:=ReplaceText(SQL, ':grorderid', VarToSQL(GlTaskKey));
    try
      Err:=ExecuteSQLCommit(SQL);
    except
      showmessage('Помилка прив''язки змінного завдання склопакетів до замовлення конструкцій'+#13+ExceptionMessage+#13+SQL);
      OnCloseProgram;
      exit;
    end;

    //-------------------------------------------------------------------------- Підрахунок кількості створених кроїв Металу
    n:=0;
    i:=0;
    while i<GlOrdersDL.Count do begin
      if GlOrdersDL[i]['isCreated']=1 then begin
        inc(n)
      end;
      inc(i);
    end;
    if UserContext.UserName='victor' then begin
      showmessage('Оброблено:'+#13+
                  '  Кроїв: '+VarToStr(GLOrdersDL.Count)+#13+
                  '  Кроїв з металом: '+VarToStr(n)
      );
    end;
    GlTask:=OpenDocument(IdocGlassTask, GlTaskKey);
    GlTask.Show;
  end;
  OnCloseProgram;
End;
