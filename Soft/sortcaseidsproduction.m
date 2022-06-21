function sortcaseidsproduction()

% Sort case IDs based on priority.
disp('Sorting out case IDs - please wait.')

%load Temp\values.mat values
load Temp\caseids_production.mat caseids_production
load Temp\onlinereport.mat onlinereport

temp = caseids_production;

% Get the customernumber
% col_accountnr = catchcolumnindex({'DeliveryOfiiceAccountNumber'},onlinereport,1);
% col_accountnr = cell2mat(col_accountnr(2,1));
% col_estshipdate = catchcolumnindex({'SurgeryDate'},onlinereport,1);
% col_estshipdate = cell2mat(col_estshipdate(2,1));
col_estshipdate = catchcolumnindex({'CaseId'},onlinereport,1); % shipping date is based on the primary key, which is in chronological order.
col_estshipdate = cell2mat(col_estshipdate(2,1));

% For all cases
for x = 1:size(temp,1)
    % Get the case ID or RS-number
    cc = temp(x,1);
    % disp(['Current case ID: ' char(cc)]);
    if size(char(cc),2) == 12
        % Get the position in the online report
        row_caseid = catchrowindex(cc,onlinereport,2);
        row_caseid = cell2mat(row_caseid(2,1));
        temp(x,2) = {num2str(row_caseid)};
        temp(x,3) = {num2str(cell2mat(onlinereport(row_caseid,col_estshipdate)))};
    elseif size(char(cc),2) == 13
        temp(x,2) = {'99999999999'};
        temp(x,3) = {'0'};
    else
        disp(char(cc));
        error('Case ID or RS-number is not valid.');
    end
end

% temp2 = sortrows(temp,[3,2]);
temp2 = sortrows(temp,1);
caseids_production = temp2(:,1);

save Temp\caseids_production.mat caseids_production

end

