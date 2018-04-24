clear all;close all;clc;
[ndata,text,alldata]=xlsread('F:\GBT14885-94','sheet1');


BH=alldata(:,11);
BH1=BH(4:8548);

s=xlswrite('F:\GBT14885-94',BH1,'sheet1','L4:L8548');
