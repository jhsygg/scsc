function YuanZi=F_RAM_450_YZ(fname1,fname2,ZX)
[ndata1,text1,alldata1]=xlsread(fname1,ZX);%读取计算数据
[ndata2,text2,alldata2]=xlsread(fname2,ZX);%读取计算数据

BH=text1(:,3);
BH2=text2(:,2);
n1=numel(BH);
n2=numel(BH2);
BH1=BH(2:n1);BH3=BH2(2:n2);
YZ=ndata2(:,1);
ZJ=ndata2(:,2);

%for i=1:(n1-1)
%    for j=1:(n2-1)
%       if strcmp(BH1(i),BH3(j))
          %S0=xlswrite(fname1,YZ(j),ZX,['R',int2str(i+1)]);
          %S1=xlswrite(fname1,ZJ(j),ZX,['W',int2str(i+1)]);
%          break;
%       end
%    end
%end

