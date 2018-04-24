clear all;close all;clc;
[ndata,text,alldata]=xlsread('F:\RAM数据\RAM-450-DZB','DZ');%读取计算数据

DW=text(:,9);
n1=numel(DW);
DW1=DW(2:n1);
BM=cell('高炉车间');
for i=1:(n1-1)
    if findstr(char(DW1(i)),'高炉')==1
        S0=xlswrite('F:\RAM数据\RAM-450-DZB',BM,'DZ',['H',int2str(i+1)];)
    end
end

%S0=xlswrite('F:\RAM数据\RAM-450-TYB',BH1,'TY','C2:C1455');
