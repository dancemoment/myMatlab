[data,text]=xlsread(['Y:\Documents\Work\TianHongFund\basicData','\ashareconseption.xlsx'],'ashareconseption');
text(1,:)=[];
dateall=datenum(datetime(data(:,[1 2]),'ConvertFrom','excel'));
dateInNum=dateall(:,1);
dateInCell=cellstr(datestr(dateInNum,'yyyy-mm-dd'));
dateOutNum=dateall(:,2);
dateOutNum(isnan(dateOutNum))=datenum(today);
dateOutCell=cellstr(datestr(dateOutNum,'yyyy-mm-dd'));
consepCode=data(:,3);

sumData(max(consepCode),1)=struct;
conceptTextU=unique(text(:,1),'stable');
fieldList='mkt_cap_float,close,pct_chg,volume,turn,mrg_long_bal,mrg_short_bal,mf_vol_ratio,xq_WOW_focus,xq_WOW_comments,pe_ttm,peg,eps_basic,roe_avg,operateincometoebt,ocftooperateincome_ttm,yoyeps_basic,yoy_or';
field={'mkt_cap_float','close','pct_chg','volume','turn','mrg_long_bal','mrg_short_bal','mf_vol_ratio','xq_WOW_focus','xq_WOW_comments','pe_ttm','peg','eps_basic','roe_avg','operateincometoebt','ocftooperateincome_ttm','yoyeps_basic','yoy_or'};
filedName={'流通市值','收盘价','涨跌幅','成交量','换手率','融资余额','融券余额','资金流向占比','雪球关注度增长率','雪球讨论增长率','ttm市盈率','peg','基本每股收益','roe','经营利润税前收益比','经营活动现金流经营利润比','eps同比增长率','同比增长率'};
% 资金流向占比：当周期大单资金净流入占流通股本的比率
for k=1:max(consepCode)
    sumData(k,1).conceptName=conceptTextU{k};
    stockListI=text(consepCode==k,2);
    sumData(k,1).stockList=stockListI;
    sumData(k,1).dateincell=dateInCell(consepCode==k);
    sumData(k,1).dateoutcell=dateOutCell(consepCode==k);
    sumData(k,1).dataField=fieldList;
    
    for i=1:length(sumData(k,1).stockList)
        [data,code,fieldwind,time]=w.wsd(sumData(k,1).stockList{i},fieldList,sumData(k,1).dateincell{i},sumData(k,1).dateoutcell{i},'currencyType=','Period=D','gRateType=1','Fill=Previous','PriceAdj=F');
        sumData(k,1).dataInfo(i,1).data=data;
        sumData(k,1).dataInfo(i,1).time=time;
        sumData(k,1).dataInfo(i,1).field=fieldwind;
    end
    %
    for j=1:length(field)
        DATA=outerjoin(table(sumData(k,1).dataInfo(1,1).time,sumData(k,1).dataInfo(1,1).data(:,j),'VariableNames',{'time',field{j}}),table(sumData(k,1).dataInfo(2,1).time,sumData(k,1).dataInfo(2,1).data(:,j),'VariableNames',{'time',field{j}}),'keys','time');
        timeSet=union(DATA.time_left,DATA.time_right);
        timeSet(isnan(timeSet))=[];
        DATA.time=timeSet;
        DATA.time_left=[];
        DATA.time_right=[];
        DATA.Properties.VariableNames{[field{j},'_left']}='stock1';
        DATA.Properties.VariableNames{[field{j},'_right']}='stock2';
    
        for m=3:length(sumData(k,1).stockList)
            DATA=outerjoin(DATA,table(sumData(k,1).dataInfo(m,1).time,sumData(k,1).dataInfo(m,1).data(:,j),'VariableNames',{'time',field{j}}),'keys','time');
            timeSet=union(DATA.time_DATA,DATA.time_right);
            timeSet(isnan(timeSet))=[];
            DATA.time=timeSet;
            DATA.time_DATA=[];
            DATA.time_right=[];
            DATA.Properties.VariableNames{field{j}}=['stock',num2str(m)];
        end
        
        sumData(k,1).mergeData.(field{j})=DATA;
    end
end
%%
for k=1:max(consepCode)
    wtsData=sumData(k,1).mergeData.(field{1});
    wtsData.time=[];
    wts=table2array(wtsData);
    wtsDenom=repmat(nansum(wts,2),1,size(wtsData,2));
    wts=wts./wtsDenom;
    for j=1:length(field)
        Data=sumData(k,1).mergeData.(field{j});
        time=Data.time;
        Data.time=[];
        numData=table2array(Data);
        sumData(k,1).([field{j},'Serial'])=table(time,nansum(numData.*wts,2),'VariableNames',{'time',field{j}});
    end
end
%% 获取主题同期hs300收益率
indexCode='000300.SH';
for k=1:length(sumData)
    beginDate=datestr(sumData(k,1).turnSerial.time(1),'yyyy-mm-dd');
    endDate=datestr(sumData(k,1).turnSerial.time(end),'yyyy-mm-dd');
    
    [indexRet,~,~,indexTime]=w.wsd(indexCode,'pct_chg',beginDate,endDate,'Fill=Previous','PriceAdj=F');

    sumData(k,1).indexRet=[indexRet,indexTime];
end
for k=1:length(sumData)
    sumData(k,1).indexRetTable=table(sumData(k,1).indexRet(:,1),sumData(k,1).indexRet(:,2),'VariableName',{'indexRet','time'});
end
%% summary info regarding different field
serialAll=struct;
for i=1:length(field)
    abInfo=outerjoin(sumData(1,1).([field{i},'Serial']),sumData(2,1).([field{i},'Serial']),'keys','time');
    timenew=union(abInfo.time_left,abInfo.time_right);
    timenew(isnan(timenew))=[];
    abInfo.time=timenew;
    abInfo.time_left=[];
    abInfo.time_right=[];
    abInfo.Properties.VariableNames{[field{i},'_left']}='conception1';
    abInfo.Properties.VariableNames{[field{i},'_right']}='conception2';
    for k=3:length(sumData)

        abInfo=outerjoin(abInfo,sumData(k,1).([field{i},'Serial']),'keys','time');
        timenew=union(abInfo.time_abInfo,abInfo.time_right);
        timenew(isnan(timenew))=[];
        abInfo.time=timenew;
        abInfo.time_abInfo=[];
        abInfo.time_right=[];
        abInfo.Properties.VariableNames{field{i}}=['conception',num2str(k)];
    end
    serialAll.(field{i})=abInfo;
end
%% generate growth data
 %每次只取异动前三名的主题x
for i=2:8
   top=3;
   datainfoTable=serialAll.(field{i});
   datainfoDouble=table2array(datainfoTable);
   datagrowth=datainfoDouble(:,1:end-1);
   zeroExist=find(datagrowth==0);
   if ~isempty(zeroExist)
       datagrowth(zeroExist)=NaN;
   end 
   datagrowth=price2ret(datagrowth);
   datagrowthInfo=[datagrowth,datainfoDouble(2:end,end)];
   growthInfo.(field{i})=datagrowthInfo;
   for t=datagrowthInfo(:,end)'
       datePos=find(datagrowthInfo(:,end)==t);
       report.(field{i}){datePos,1}=t;
       nn=isnan(datagrowth(datePos,:));
       datagrowth(datePos,nn)=-100;
       [info,idx]=sort(datagrowth(datePos,1:end-1),'descend');
       info=info(1:top);
       idx=idx(1:top);
       concept=conceptTextU';
       report.(field{i}){datePos,2}=concept(idx);
       report.(field{i}){datePos,3}=info;
       for j=1:length(idx)
           arrayData=table2array(sumData(idx(j)).mergeData.(field{i}));
           date2Pos=find(arrayData(:,end)==t);
           if isempty(date2Pos)
               report.(field{i}){datePos,j+3}=nan;
           else
               specInfo=arrayData(date2Pos,1:end-1);
               nnn=isnan(specInfo);
               specInfo(nnn)=-100;
               [specficInfo,specficIdx]=sort(specInfo,'descend');
               sepcficInfo=specficInfo(1:top);
               specficIdx=specficIdx(1:top);
               report.(field{i}){datePos,j+3} =sumData(idx(j)).stockList(specficIdx);
           end
       end
   end
end    
%% statistics
infoReport

   
   abInfoSum.time=str2num(datestr(abInfoSum.time,'yyyymmdd'));
   writetable(abInfoSum,fileNameAb,'sheet',sheetNameAb);
   xlswrite(fileNameAb,conceptTextU',sheetNameAb,'B1');
   fprintf('%s数据输出完毕\n',field{i});
end

%% output in xlsx
filePath=[pwd,'\heatData'];
if ~exist(filePath,'dir')
   mkdir(filePath)
end
% write abnormal return info into excel
fileNameAb=[filePath,'\heatData.xlsx'];
for i=2:8
   sheetNameAb=[field{i}];
   abInfoSum=serialAll.(field{i})(:,[end,1:end-1]);
   abInfoSum.time=str2num(datestr(abInfoSum.time,'yyyymmdd'));
   writetable(abInfoSum,fileNameAb,'sheet',sheetNameAb);
   xlswrite(fileNameAb,conceptTextU',sheetNameAb,'B1');
   fprintf('%s数据输出完毕\n',field{i});
end
      