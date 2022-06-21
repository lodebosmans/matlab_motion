function [currentprice,discount,comment] = getregularprice(salesordercheck_v3,agr_cus,agr_price,pricingmethod,flow,type2,cy,cm)

if pricingmethod.customprice == 1 || pricingmethod.toplayerprice
    [discountcascade] = checkfordiscount(agr_price,pricingmethod.customprice,pricingmethod.toplayerprice);
end

if strcmp(upper(agr_price.SplitTypeCounter),'YES') == 1
    stc = 1;
else
    stc = 0;
end

% if flow == 1
%     % Own pricing strategy
%

if pricingmethod.standardprice == 1
    
    currentprice = agr_price.(type2);
    discount = '-';
    comment = 'Standard pricing';
    
elseif pricingmethod.customprice == 1
    nroflevels = size(discountcascade.([type2 'Cus']),1);
    if nroflevels > 1
        % Compare with the counter to get the correct level from the discountcascade
        if stc == 1
            % Get the type total for that year.
            alreadyordered = size(salesordercheck_v3.orders.(agr_price.CustomerNumber).(['y' num2str(cy)]).(['Type_' type2]),1);
        else
            % Get the year total for that year.
            alreadyordered = size(salesordercheck_v3.orders.(agr_price.CustomerNumber).(['y' num2str(cy)]).overview,1);
        end
        newcase = alreadyordered + 1;
        goon = 1;
        clvl = 1;
        % Get the correct pricing level
        while goon == 1 && clvl < nroflevels
            if newcase < discountcascade.([type2 'Cus'])(clvl+1,1)
                goon = 0;
            else
                clvl = clvl + 1;
            end
        end
        currentprice = discountcascade.([type2 'Cus'])(clvl,2);
        discount = '-';
        comment = ['Level ' num2str(clvl) '/' num2str(nroflevels) ' custom pricing'];
    else
        currentprice = discountcascade.([type2 'Cus'])(1,2);
        discount = '-';
        comment = 'Level 1/1 custom pricing';
    end
elseif pricingmethod.toplayerprice == 1
    % dan wat?
    test = 1;
    
    nroflevels = size(discountcascade.(type2),1);
    if nroflevels > 1
        % Compare with the counter to get the correct level from the discountcascade
        if stc == 1
            % Get the type total for that year.
            alreadyordered = size(salesordercheck_v3.orders.(agr_price.CustomerNumber).(['y' num2str(cy)]).(['Type_' type2]),1);
        else
            % Get the year total for that year.
            alreadyordered = size(salesordercheck_v3.orders.(agr_price.CustomerNumber).(['y' num2str(cy)]).overview,1);
        end
        newcase = alreadyordered + 1;
        goon = 1;
        clvl = 1;
        % Get the correct pricing level
        while goon == 1 && clvl < nroflevels
            if newcase < discountcascade.(type2)(clvl+1,1)
                goon = 0;
            else
                clvl = clvl + 1;
            end
        end
        currentprice = discountcascade.(type2)(clvl,2);
        discount = '-';
        comment = ['Level ' num2str(clvl) '/' num2str(nroflevels) ' top layer pricing'];
    else
        currentprice = discountcascade.(type2)(1,2);
        discount = '-';
        comment = 'Level 1/1 top layer pricing';
    end
    
end

%
% elseif flow == 2
%     % Main pricing strategy
%
% elseif flow == 3 || flow == 5
%     % Agent own pricing strategy
%
% elseif flow == 4 || flow == 6
%     % Agent main pricing strategy
%
%     %     elseif flow == 5
%     %         % Agent own pricing strategy
%     %
%     %     elseif flow == 6
%     %         % Agent main pricing strategy
%
% else
%     disp(['No pricing flow has been determined. Please check ' ccid ' for customer ' cusnr ' on date ' num2str(cy) '/' num2str(cm) '/' num2str(cd) '.'])
% end


end