function [salesordercheck_v3] = checkfororderstruct(salesordercheck_v3,cusnr,cy)

if isfield(salesordercheck_v3.orders,cusnr) == 0
    % If not, created it for this year
    salesordercheck_v3.orders.(cusnr).totalordersever = 0;
    [salesordercheck_v3] = createorderstruct(salesordercheck_v3,cusnr,cy);
    % Put the fields in alphabetical order
    salesordercheck_v3.orders = orderfields(salesordercheck_v3.orders);
elseif isfield(salesordercheck_v3.orders.(cusnr),['y' num2str(cy)]) == 0
    % Check if they have ordered already this year. If not create the struct.
    [salesordercheck_v3] = createorderstruct(salesordercheck_v3,cusnr,cy);
else
    
end