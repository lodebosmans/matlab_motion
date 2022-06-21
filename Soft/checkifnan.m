function [database] = checkifnan(name,currentcaseid,database,onlinereport,rownr_OR,colnr_OR,rownr_overview,colnr_overview)

% Check if NaN
value = onlinereport(rownr_OR,colnr_OR);

%     % Check if character
%     if iscellstr(onlinereport(rownr_OR,colnr_OR))
%         database.raw.(currentcaseid).(name) = char(onlinereport(rownr_OR,colnr_OR));
%         database.overview(rownr_overview,colnr_overview) = cellstr(database.raw.(currentcaseid).(name));
%         % Check if number
%     elseif isnumeric(cell2mat(onlinereport(rownr_OR,colnr_OR)))
%         database.raw.(currentcaseid).(name) = cell2mat(onlinereport(rownr_OR,colnr_OR));
%         database.overview(rownr_overview,colnr_overview) = num2cell(database.raw.(currentcaseid).(name));
%     end

if isa(value{1},'double') == 1
    database.raw.(currentcaseid).(name) = num2str(cell2mat(value));
    database.overview(rownr_overview,colnr_overview) = cellstr(database.raw.(currentcaseid).(name));
elseif isa(value{1},'char') == 1
    database.raw.(currentcaseid).(name) = char(value);
    database.overview(rownr_overview,colnr_overview) = cellstr(database.raw.(currentcaseid).(name));
else
    database.raw.(currentcaseid).(name) = '';
end

end

