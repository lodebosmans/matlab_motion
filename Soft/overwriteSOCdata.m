% Does not work
for i = 1:size(salesordercheck_v3.data,1)
    for j = 1:size(salesordercheck_v3.data,2)
        if i == 18560
            stop = 1;
        end
        if isempty(salesordercheck_v3.data(i,j))
            salesordercheck_v3.data(i,j) = {'-'};
        end
    end
end


% Works
salesordercheck_v3.data(cellfun(@(x) isempty(x), salesordercheck_v3.data)) = {'-'};