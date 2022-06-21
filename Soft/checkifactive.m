function [active] = checkifactive(startyear,startmonth,startday,cy,cm,cd,endyear,endmonth,endday)

% Check if the condition is still ongoing
if cy == startyear && cm == startmonth && cd >= startday || ...
        cy == startyear && cm > startmonth || ...
        cy > startyear
    % The current agreement is valid based on the starting date
    % Check if the end date hasn't passed yet.
    if cy == endyear && cm == endmonth && cd > endday || ...
            cy == endyear && cm > endmonth || ...
            cy > endyear
        % The current agreement is not valid based on the ending date
        active = 0;
    else
        % The current agreement is  valid based on the ending date
        active = 1;
    end    
else
    % The agreement hasn't started yet.
    active = 0;
end

end