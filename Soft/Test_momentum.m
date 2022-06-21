
for x = 1:size(salesordercheck_v3.data,1)
    if strcmp(salesordercheck_v3.data(x,3),'Funct IntellTraining Company') == 1
        salesordercheck_v3.data(x,3) = {'Momentum Sports Injury Company'};
        salesordercheck_v3.data(x,4) = {'Momentum Sports Injury'};
    end
end