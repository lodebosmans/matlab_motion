function checkomsstatus()

disp(' ');
disp('----------------------------------------');
disp('Checking the OMS statusses - please wait');
disp('----------------------------------------');
disp(' ');

load Temp\values.mat values
load Temp\onlinereport.mat onlinereport

col_status = catchcolumnindex({'CaseStatus'},onlinereport,1);
col_status = cell2mat(col_status(2,1));
failure = 0;
failedcases = cell(0);

% Get the case IDs for that month
[caseidsmonth] = getcaseidsforspecificmonth();

nrofcases = size(caseidsmonth,1);
for crow = 1:nrofcases
    % Get the case ID
    ccid = char(caseidsmonth(crow,1));  
    disp(ccid);
    % Get the position in the online report
    row_ccid = catchrowindex({ccid},onlinereport,2);
    row_ccid = cell2mat(row_ccid(2,1));    
    % Get the status for this case and compare with shipped
    if strcmp(onlinereport(row_ccid,col_status),'Shipped') == 1
        % Do nothing    
    else
%         if failure == 0
%             %disp(['Cases that are not in status shipped for ' values.RequestedMonth '/' values.RequestedYear ':']);
%         end
        failure = failure + 1;
        disp([ccid ' failed.']);
        nextone = size(failedcases,1) + 1;
        failedcases(nextone,1) = {ccid};
    end
end

if failure == 0
    disp(' ');
    disp(['Hooray! All cases for ' values.RequestedMonth '/' values.RequestedYear ' are in status shipped.']);
else
    disp(' ');
    disp(['Not all cases for ' values.RequestedMonth '/' values.RequestedYear ' are in status shipped.']);
    disp('Failed cases: ')
    for x = 1:failure
        ccid = char(failedcases(x,1));
        disp(ccid);
    end
end
disp(['Success rate = ' num2str(nrofcases - failure) '/' num2str(nrofcases)]);

end