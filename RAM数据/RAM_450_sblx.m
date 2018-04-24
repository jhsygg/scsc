clear all;close all;clc;
[ndata,text,alldata]=xlsread('F:\RAM数据\GBT14885-94D','GB');%读取计算数据

DH1=text(:,19);DH2=text(:,20);DH3=text(:,21);
n1=numel(DH1);

S0=xlswrite('F:\RAM数据\设备类别表',DH1,'设备类别','A2,A2245');
S1=xlswrite('F:\RAM数据\设备类别表',DH2,'设备类别','B2:B2245');
S2=xlswrite('F:\RAM数据\设备类别表',DH3,'设备类别','C2:C2245');
S4=xlswrite('F:\RAM数据\设备类别表',DH1,'设备类别','D2:D2245');

