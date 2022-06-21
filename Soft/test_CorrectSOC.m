SOCrows = size(salesordercheck_v3.data,1);
for x = 1:SOCrows
   if strcmp(char(salesordercheck_v3.data(x,5)),'0: Cancelled') == 1
       salesordercheck_v3.data(x,2) = {'-'};
       salesordercheck_v3.data(x,3) = {'-'};
       salesordercheck_v3.data(x,4) = {'-'};
       salesordercheck_v3.data(x,6) = {'-'};
       salesordercheck_v3.data(x,7) = {'-'};
       salesordercheck_v3.data(x,8) = {'-'};
       salesordercheck_v3.data(x,9) = {'-'};
       salesordercheck_v3.data(x,10) = {'-'};
       salesordercheck_v3.data(x,11) = {'-'};
       salesordercheck_v3.data(x,12) = {'-'};
       salesordercheck_v3.data(x,13) = {'-'};
       salesordercheck_v3.data(x,14) = {'-'};
       salesordercheck_v3.data(x,15) = {'-'};
       salesordercheck_v3.data(x,16) = {'-'};
       salesordercheck_v3.data(x,17) = {'-'};
       
   end
    
end