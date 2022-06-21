
[a,b,c] = xlsread('C:\Users\mathlab\Documents\Matlab\20190318_Interface\CasesFB.xlsx');

result = cell(0);

for x = 1:size(c,1)
    caseid = char(c(x,1));
    disp(caseid)
    
%     find(caseid,salesordercheck_v3.datafeed(:,1))
    
    idx = find(contains(onlinereport(:,2),caseid));
    disp(idx);
    
    distributor = onlinereport(idx,22);
    country = onlinereport(idx,19);
    
    result(x,1) = cellstr(caseid);
    result(x,2) = distributor;
    result(x,3) = country;
    
end