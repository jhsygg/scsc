clear all;close all;clc;
[ndata1,text1,alldata1]=xlsread('F:\RAM����\RAM-450-HZB','HZB');%��ȡ��������
[ndata2,text2,alldata2]=xlsread('F:\RAM����\�豸����','�豸����');%��ȡ��������

BH=text2(:,2);
n1=numel(BH);
BH2=text1(:,25);
n2=numel(BH2);
BH1=BH(2:n1);
BH3=BH2(2:n2);
ZCM=text2(:,3);
ZCM1=ZCM(2:n1);
ZCM2=text1(:,2);

for i=1:(n2-1)
    for j=1:(n1-1)
       if strcmp(BH3(i),BH1(j))
           ZCM2(i)=ZCM1(j);           
           break;
       end
    end
end
S0=xlswrite('F:\RAM����\RAM-450-HZB',ZCM2,'HZB','G2:G2001');


