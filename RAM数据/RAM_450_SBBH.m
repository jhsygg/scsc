clear all;close all;clc;
[ndata,text,alldata]=xlsread('F:\RAM数据\RAM-450-HZB','HZB');%读取计算数据

DH=text(:,3);
n1=numel(DH);
BH1=DH(2:n1);




%S0=xlswrite('F:\RAM数据\RAM-450-DZB',BH1,'DZ','C2:C50');
%s2=xlswrite('F:\RAM数据\RAM-450-DZB',Date1,'DZ','K2:K50');
%s3=xlswrite('F:\RAM数据\RAM-450-DZB',Date1,'DZ','N2:N50');
%s4=xlswrite('F:\RAM数据\RAM-450-DZB',Date1,'DZ','O2:O50');
%s5=xlswrite('F:\RAM数据\RAM-450-DZB',Azdd1,'DZ','I2:I50');
%s6=xlswrite('F:\RAM数据\RAM-450-DZB',Zzc1,'DZ','J2:J50');
%s7=xlswrite('F:\RAM数据\RAM-450-DZB',Sum1,'DZ','AD2:AD50');
%s8=xlswrite('F:\RAM数据\RAM-450-DZB',Name1,'DZ','D2:D50');
%s9=xlswrite('F:\RAM数据\RAM-450-DZB',Xh1,'DZ','E2:E50');
%s10=xlswrite('F:\RAM数据\RAM-450-DZB',Yz1,'DZ','R2:R50');
%S1=xlswrite('F:\RAM数据\RAM-450-DZB',Date3,'DZ','X2:X50');



