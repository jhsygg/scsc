clear all;close all;clc;
[ndata1,text1,alldata1]=xlsread('F:\RAM数据\部门设置表','部门设置表','b2:c29');%读取计算数据
[ndata2,text2,alldata2]=xlsread('F:\RAM数据\RAM-450-HZB1','HZB','h2:h2001');%读取计算数据
[ndata3,text3,alldata3]=xlsread('F:\RAM数据\RAM-450-HZB1','HZB');%读取计算数据

%生成空表

bh1=numel(text2);
%BH=cell(bh1,1);
ZCJH=cell(bh1,5);
RQ=cell(bh1,1);
LX=zeros(bh1,1);
SL=zeros(bh1,1);
YZ=zeros(bh1,1);
ZJ=zeros(bh1,1);

bh2=numel(text1)/2;

%生成部门代号
for i=1:bh1
    ZCJH(i,1)=text3(i+1,3);
    ZCJH(i,3)=text3(i+1,8);
    ZCJH(i,4)=text3(i+1,4);
    ZCJH(i,5)=text3(i+1,5);
    RQ(i,1)=text3(i+1,11);
    LX(i)=ndata3(i,15);
    SL(i)=ndata3(i,35);
    YZ(i)=ndata3(i,14);
    ZJ(i)=ndata3(i,19);
    for j=1:(bh2)
        if strcmp(text2(i,1),text1(j,2))
            ZCJH(i,2)=text1(j,1);
         end
    end
end

S0=xlswrite('F:\RAM数据\数据交换模板',ZCJH,'数据交换模板','B2:F2001');
S11=xlswrite('F:\RAM数据\数据交换模板',SL,'数据交换模板','I2:I2001');
S10=xlswrite('F:\RAM数据\数据交换模板',RQ,'数据交换模板','J2:J2001');
S1=xlswrite('F:\RAM数据\数据交换模板',LX,'数据交换模板','K2:K2001');
S2=xlswrite('F:\RAM数据\数据交换模板',SL,'数据交换模板','I2:I2001');
S3=xlswrite('F:\RAM数据\数据交换模板',YZ,'数据交换模板','L2:L2001');
S4=xlswrite('F:\RAM数据\数据交换模板',ZJ,'数据交换模板','M2:M2001');

