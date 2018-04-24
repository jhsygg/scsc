clear all;close all;clc;
[ndata,text,alldata]=xlsread('F:\RAM数据\M-450','TY');%读取计算数据

DH=text(:,1);
n1=numel(DH);
DH1=DH(2:n1);
Name=text(:,10);
Name1=Name(2:n1);%资产名称
Xh=alldata(:,11);
Xh1=Xh(2:n1);%规格型号
BH=text(:,23);
BH1=BH(2:n1);%固定资产号
Date=text(:,24);
Date1=Date(2:n1);%形成日期
Azdd=text(:,19);
Azdd1=Azdd(2:n1);%安装地点
Sum=alldata(:,17);
Sum1=Sum(2:n1);%数量
Zzc=text(:,12);
Zzc1=Zzc(2:n1);%制造厂家
Yz=alldata(:,18);
Yz1=Yz(2:n1);%原值
Date2=text(:,25);
Date3=Date2(2:n1);

S0=xlswrite('F:\RAM数据\RAM-450-TYB',BH1,'TY','C2:C1455');
s2=xlswrite('F:\RAM数据\RAM-450-TYB',Date1,'TY','K2:K1455');
s3=xlswrite('F:\RAM数据\RAM-450-TYB',Date1,'TY','N2:N1455');
s4=xlswrite('F:\RAM数据\RAM-450-TYB',Date1,'TY','O2:O1455');
s5=xlswrite('F:\RAM数据\RAM-450-TYB',Azdd1,'TY','I2:I1455');
s6=xlswrite('F:\RAM数据\RAM-450-TYB',Zzc1,'TY','J2:J1455');
s7=xlswrite('F:\RAM数据\RAM-450-TYB',Sum1,'TY','AD2:AD1455');
s8=xlswrite('F:\RAM数据\RAM-450-TYB',Name1,'TY','D2:D1455');
s9=xlswrite('F:\RAM数据\RAM-450-TYB',Xh1,'TY','E2:E1455');
s10=xlswrite('F:\RAM数据\RAM-450-TYB',Yz1,'TY','R2:R1455');
S1=xlswrite('F:\RAM数据\RAM-450-TYB',Date3,'TY','X2:X1455');



