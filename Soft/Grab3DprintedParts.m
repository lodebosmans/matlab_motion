clear all
close all
clc

year = 2019;
month = 2;

%%

year2 = num2str(year);
if month < 10 
    month2 = ['0' num2str(month)];
else
    month2 = num2str(month);
end

load C:\Users\Lode\Documents\Matlab\RSPrint\Interface\Input\salesordercheck_v3.mat salesordercheck_v3
[a,b,raw] = xlsread('C:\Users\Lode\Documents\Matlab\RSPrint\20190228_ShipmentMaterialiseBe_Combined RSPRINT_Go4D_Livit_Rebuilts_US_finaal',1);
success = 0;

for x = 1:size(raw,1)
    ccid = raw(x,1);
    
    if strcmp(char(ccid),'RS18-JAZ-NIH') == 1
        %skip        
    else
        [input] = catchrowindex(ccid,salesordercheck_v3.data,1);
        rownr = cell2mat(input(2,1));
        
        itemnumber = char(salesordercheck_v3.data(rownr,6));
        if strcmp(itemnumber(2),'X') == 1
            disp(['Error with ' char(ccid)]);
            raw(x,2) = {['Temp itemnumber ' itemnumber]};
        else
            raw(x,2) = {['3D' itemnumber(2:3)]};
            success = success + 1;
        end        
    end    
end
disp(['Success = ' num2str(success) '/' num2str(x)])

xlswrite('C:\Users\Lode\Documents\Matlab\RSPrint\Interface_TestEnvironment2_pricing\Output\CaseIDs\caseIDs.xlsx',raw)

disp('Script finished');