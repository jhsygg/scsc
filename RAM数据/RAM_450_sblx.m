clear all;close all;clc;
[ndata,text,alldata]=xlsread('F:\RAM����\GBT14885-94D','GB');%��ȡ��������

DH1=text(:,19);DH2=text(:,20);DH3=text(:,21);
n1=numel(DH1);

S0=xlswrite('F:\RAM����\�豸����',DH1,'�豸���','A2,A2245');
S1=xlswrite('F:\RAM����\�豸����',DH2,'�豸���','B2:B2245');
S2=xlswrite('F:\RAM����\�豸����',DH3,'�豸���','C2:C2245');
S4=xlswrite('F:\RAM����\�豸����',DH1,'�豸���','D2:D2245');

