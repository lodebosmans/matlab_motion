function [endyear,endmonth,active,period,currentmonth] = getperiod(values,startyear,startmonth,duration)

% See how many full years are included in the duration
nroffullyears = floor(duration/12);
% Calculate the remaining months
nrofremainingmonths = duration - nroffullyears*12;
% Check if you pass Lode's birthday (1st of January) by adding the remaining months to the start month
if startmonth + nrofremainingmonths > 12
    addextrayear = 1;
    endmonth = startmonth + nrofremainingmonths - 12 -1; % -1 to avoid double counting of the first month
else
    addextrayear = 0;
    endmonth = startmonth + nrofremainingmonths - 1; % -1 to avoid double counting of the first month
end
% Define the end date
endyear = startyear + nroffullyears + addextrayear;
% Check if the condition is still ongoing
active = 0;
if values.MenuYear_nr < endyear
    active = 1;
elseif values.MenuYear_nr == endyear
    if values.MenuMonth_nr <= endmonth
        active = 1;
    end
else
    active = 0;
end
% Calculate the current month of the period if the condition is still active
period = zeros(4,duration);
if active == 1
    month = startmonth - 1;
    year = startyear;
    for y = 1:duration
        period(1,y) = y;
        month = month + 1;
        if month > 12
            month = month - 12;
            year = year + 1;
        end
        period(3,y) = month;
        period(2,y) = year;
        if values.MenuYear_nr == period(2,y) && values.MenuMonth_nr == period(3,y)
            period(4,y) = 1;
        else
            period(4,y) = 0;
        end
    end
    currentmonth = find(period(4,:)==1);
    if isempty(currentmonth) == 1
        currentmonth = 0;
    end
end


end