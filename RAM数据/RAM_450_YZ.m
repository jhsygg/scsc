clear all;close all;clc;
[ndata1,text1,alldata1]=xlsread('F:\RAM数据\RAM-450-HZB1','HZB');%读取计算数据
[ndata2,text2,alldata2]=xlsread('F:\RAM数据\2011-12-YYB','HZB');%读取计算数据

BH=text1(:,3);
BH2=text2(:,2);
n1=numel(BH);
n2=numel(BH2);
BH1=BH(2:n1);BH3=BH2(2:n2);CC=ndata1(:,16);YY=ndata1(:,16);
YZ=ndata2(:,14);
ZJ=ndata2(:,15);
n3=numel(CC);
CC1=CC(2:n3);
YY1=YY(2:n3);

for i=1:(n1-1)
    for j=1:(n2-1)
       if strcmp(BH1(i),BH3(j))
           CC1(i)=YZ(j);
           YY1(i)=ZJ(j);
         break;
       end
    end
end
S0=xlswrite('F:\RAM数据\RAM-450-HZB1',CC1,'HZB','R2:R2001');
S1=xlswrite('F:\RAM数据\RAM-450-HZB1',YY1,'HZB','W2:W2001');