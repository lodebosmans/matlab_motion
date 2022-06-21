function [discountcascade] = checkfordiscount(agr_price,customprice,toplayerprice)

% Get the correct itemlabels
if customprice == 1
    items = {'RegSemiCus','RegFullCus','SlimSemiCus','SlimFullCus'};
elseif toplayerprice == 1
    items = {'TL_RegSemi_None','TL_RegSemi_Eva_NotAs','TL_RegSemi_Eva_As','TL_RegSemi_Pusoft_NotAs','TL_RegSemi_PuSoft_As','TL_RegSemi_EvaCarbon_NotAs','TL_RegSemi_EvaCarbon_As',...
             'TL_RegFull_None','TL_RegFull_Eva_NotAs','TL_RegFull_Eva_As','TL_RegFull_Pusoft_NotAs','TL_RegFull_PuSoft_As','TL_RegFull_EvaCarbon_NotAs','TL_RegFull_EvaCarbon_As',...
             'TL_SlimSemi_None','TL_SlimSemi_Eva_NotAs','TL_SlimSemi_Eva_As','TL_SlimSemi_Pusoft_NotAs','TL_SlimSemi_PuSoft_As','TL_SlimSemi_EvaCarbon_NotAs','TL_SlimSemi_EvaCarbon_As',...
             'TL_SlimFull_None','TL_SlimFull_Eva_NotAs','TL_SlimFull_Eva_As','TL_SlimFull_Pusoft_NotAs','TL_SlimFull_PuSoft_As','TL_SlimFull_EvaCarbon_NotAs','TL_SlimFull_EvaCarbon_As'};
end

% List the different options
nrofitems = size(items,2);
for x = 1:nrofitems
    ci = char(items(1,x));
    % Create the cascade cell
    eval(['discountcascade.' ci ' = [];']);
    eval(['signequal = strfind(char(agr_price.' ci '),''='');']);
    eval(['signslash = strfind(char(agr_price.' ci '),''/'');']);
    nrlevels = size(signequal,2);
    if eval(['isnumeric(agr_price.' ci ') == 1;'])
        eval(['discountcascade.' ci '(1,1) = 1;']);
        eval(['discountcascade.' ci '(1,2) = agr_price.' ci ';']);
    elseif nrlevels > 1
        for y = 1:nrlevels
            %             if y == 1
            %                 discountcascade.RegSemiCus(y,1) = str2double(agr_price.RegSemiCus(1,signequal(1,y)-1));
            %                 discountcascade.RegSemiCus(y,2) = str2double(agr_price.RegSemiCus(1,signequal(1,y)+1:signslash(1,y)-1));
            %             elseif y > 1 && y < nrlevels
            %                 discountcascade.RegSemiCus(y,1) = str2double(agr_price.RegSemiCus(1,signslash(1,y-1)+1:signequal(1,y)-1));
            %                 discountcascade.RegSemiCus(y,2) = str2double(agr_price.RegSemiCus(1,signequal(1,y)+1:signslash(1,y)-1));
            %             elseif y == nrlevels
            %                 discountcascade.RegSemiCus(y,1) = str2double(agr_price.RegSemiCus(1,signslash(1,y-1)+1:signequal(1,y)-1));
            %                 discountcascade.RegSemiCus(y,2) = str2double(agr_price.RegSemiCus(1,signequal(1,y)+1:end));
            %             end
            if y == 1
                columnamount = 'signequal(1,y)-1';
                columnprice = 'signequal(1,y)+1:signslash(1,y)-1';
            elseif y > 1 && y < nrlevels
                columnamount = 'signslash(1,y-1)+1:signequal(1,y)-1';
                columnprice = 'signequal(1,y)+1:signslash(1,y)-1';
            elseif y == nrlevels
                columnamount = 'signslash(1,y-1)+1:signequal(1,y)-1';
                columnprice = 'signequal(1,y)+1:end';
            end            
            eval(['discountcascade.' ci '(y,1) = str2double(agr_price.' ci '(1,' columnamount '));']);
            eval(['discountcascade.' ci '(y,2) = str2double(agr_price.' ci '(1,' columnprice '));']);
        end
    end
end

end