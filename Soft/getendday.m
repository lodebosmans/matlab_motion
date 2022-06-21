function [endday] = getendday(endyear,endmonth,endday)

if endmonth == 1 || endmonth == 3 || endmonth == 5 || endmonth == 7 || endmonth == 8 || endmonth == 10 || endmonth == 12
    endday = 31;
elseif endmonth == 4 || endmonth == 6 || endmonth == 9 || endmonth == 11
    endday = 30;
elseif endmonth == 2
    if floor(endyear/4) == endyear/4
        % Schrikkeljaar
        endday = 29;
    else
        endday= 28;
    end
end

end