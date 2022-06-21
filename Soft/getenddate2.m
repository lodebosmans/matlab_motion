function [endyear,endmonth,endday] = getenddate2(startyear,startmonth,startday,duration)

endday = startday - 1;

if startmonth + duration <= 12
    endyear = startyear;
    endmonth = startmonth + duration;
    if endday == 0
        [endday] = getendday(endyear,endmonth,endday);
    end
else
    nroffullyears = floor(duration/12);
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
    if endday == 0
        [endday] = getendday(endyear,endmonth,endday);
    end
end

end