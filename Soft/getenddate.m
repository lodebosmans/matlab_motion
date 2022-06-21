function [endyear,endmonth,active,currentmonth,period] = getenddate(values,startyear,startmonth,duration)

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
if active == 1
    month = startmonth;
    year = startyear;
    for y = 1:duration
        period(1,y) = y;
        month = startmonth + (y-1);
        nrofyears = floor(month/12);
        if nrofyears > 0
            month = month - nrofyears*12;
            year = year + nrofyears;
        end
   
        period(3,y) = month;
        period(2,y) = year;
    end
    
    if startyear == values.MenuYear_nr 
        values.MenuMonth_nr - startmonth + 1; % Take the first month into account
    end
    period = 1;
end

end