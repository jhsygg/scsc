clear all;close all;clc;
[ndata1,text1,alldata1]=xlsread('F:\RAM����\�������ñ�','�������ñ�','b2:c29');%��ȡ��������
[ndata2,text2,alldata2]=xlsread('F:\RAM����\RAM-450-HZB1','HZB','h2:h2001');%��ȡ��������
[ndata3,text3,alldata3]=xlsread('F:\RAM����\RAM-450-HZB1','HZB');%��ȡ��������

%���ɿձ�

bh1=numel(text2);
%BH=cell(bh1,1);
ZCJH=cell(bh1,5);
RQ=cell(bh1,1);
LX=zeros(bh1,1);
SL=zeros(bh1,1);
YZ=zeros(bh1,1);
ZJ=zeros(bh1,1);

bh2=numel(text1)/2;

%���ɲ��Ŵ���
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

S0=xlswrite('F:\RAM����\���ݽ���ģ��',ZCJH,'���ݽ���ģ��','B2:F2001');
S11=xlswrite('F:\RAM����\���ݽ���ģ��',SL,'���ݽ���ģ��','I2:I2001');
S10=xlswrite('F:\RAM����\���ݽ���ģ��',RQ,'���ݽ���ģ��','J2:J2001');
S1=xlswrite('F:\RAM����\���ݽ���ģ��',LX,'���ݽ���ģ��','K2:K2001');
S2=xlswrite('F:\RAM����\���ݽ���ģ��',SL,'���ݽ���ģ��','I2:I2001');
S3=xlswrite('F:\RAM����\���ݽ���ģ��',YZ,'���ݽ���ģ��','L2:L2001');
S4=xlswrite('F:\RAM����\���ݽ���ģ��',ZJ,'���ݽ���ģ��','M2:M2001');

