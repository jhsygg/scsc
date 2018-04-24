clear all;close all;clc;
[ndata1,text1,alldata1]=xlsread('F:\RAM数据\RAM-450-YSB','YS');%读取计算数据
[ndata2,text2,alldata2]=xlsread('F:\RAM数据\2011-12-YYB','YS');%读取计算数据

BH=text1(:,3);
BH2=text2(:,2);
n1=numel(BH);
n2=numel(BH2);
BH1=BH(2:n1);BH3=BH2(2:n2);

for i=1:(n1-1)
    for j=1:(n2-1)
       if strcmp(BH1(i),BH3(j))
           S0=xlswrite('F:\RAM数据\RAM-450-YSB','1','YS',['B',int2str(i+1)]);
           break;
       end
    end
end

