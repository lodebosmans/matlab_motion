function senddatatorsscanportal()

% Send the necessary files to the RS Scan platform.
% Manipulate data if necessary

clc
disp('Sending the order data to the RS Scan portal')

load Temp\values.mat values
if isfile('Temp\salesordercheck_v3.mat') == 0
    copyfile([values.salesordercheck_folder 'salesordercheck_v3.mat'],'Temp\salesordercheck_v3.mat'); % always keep full pathway
end
load Temp\salesordercheck_v3.mat salesordercheck_v3
load Temp\onlinereport.mat onlinereport

col_df_caseid = catchcolumnindex({'CaseCode'},salesordercheck_v3.datafeed,1);
col_df_caseid = cell2mat(col_df_caseid(2,1));
col_df_built = catchcolumnindex({'Built'},salesordercheck_v3.datafeed,1);
col_df_built = cell2mat(col_df_built(2,1));
col_df_production = catchcolumnindex({'Production'},salesordercheck_v3.datafeed,1);
col_df_production = cell2mat(col_df_production(2,1));
col_df_readytoship = catchcolumnindex({'ReadyToShip'},salesordercheck_v3.datafeed,1);
col_df_readytoship = cell2mat(col_df_readytoship(2,1));
col_df_shipped = catchcolumnindex({'Shipped'},salesordercheck_v3.datafeed,1);
col_df_shipped = cell2mat(col_df_shipped(2,1));
col_df_shippedyear = catchcolumnindex({'ShippedYear'},salesordercheck_v3.datafeed,1);
col_df_shippedyear = cell2mat(col_df_shippedyear(2,1));
col_df_shippedmonth = catchcolumnindex({'ShippedMonth'},salesordercheck_v3.datafeed,1);
col_df_shippedmonth = cell2mat(col_df_shippedmonth(2,1));
col_df_shippedday = catchcolumnindex({'ShippedDay'},salesordercheck_v3.datafeed,1);
col_df_shippedday = cell2mat(col_df_shippedday(2,1));
col_df_casestatus = catchcolumnindex({'CaseStatus'},salesordercheck_v3.datafeed,1);
col_df_casestatus = cell2mat(col_df_casestatus(2,1));
col_df_casecancelled = catchcolumnindex({'CaseCanceled'},salesordercheck_v3.datafeed,1);
col_df_casecancelled = cell2mat(col_df_casecancelled(2,1));
col_df_hopsitalaccountnumber = catchcolumnindex({'HospitalAccountNumber'},salesordercheck_v3.datafeed,1);
col_df_hopsitalaccountnumber = cell2mat(col_df_hopsitalaccountnumber(2,1));

col_or_caseid = catchcolumnindex({'CaseCode'},onlinereport,1);
col_or_caseid = cell2mat(col_or_caseid(2,1));
col_or_built = catchcolumnindex({'Built'},onlinereport,1);
col_or_built = cell2mat(col_or_built(2,1));
col_or_production = catchcolumnindex({'Production'},onlinereport,1);
col_or_production = cell2mat(col_or_production(2,1));
col_or_readytoship = catchcolumnindex({'ReadyToShip'},onlinereport,1);
col_or_readytoship = cell2mat(col_or_readytoship(2,1));
col_or_shipped = catchcolumnindex({'Shipped'},onlinereport,1);
col_or_shipped = cell2mat(col_or_shipped(2,1));
col_or_shippedyear = catchcolumnindex({'ShippedYear'},onlinereport,1);
col_or_shippedyear = cell2mat(col_or_shippedyear(2,1));
col_or_shippedmonth = catchcolumnindex({'ShippedMonth'},onlinereport,1);
col_or_shippedmonth = cell2mat(col_or_shippedmonth(2,1));
col_or_shippedday = catchcolumnindex({'ShippedDay'},onlinereport,1);
col_or_shippedday = cell2mat(col_or_shippedday(2,1));
col_or_casestatus = catchcolumnindex({'CaseStatus'},onlinereport,1);
col_or_casestatus = cell2mat(col_or_casestatus(2,1));
col_or_casecancelled = catchcolumnindex({'CaseCanceled'},onlinereport,1);
col_or_casecancelled = cell2mat(col_or_casecancelled(2,1));

% xlswrite(['datafeed.xlsx'],salesordercheck_v3.datafeed);

% Update the datafeed with all dates
for cr = 30000:size(salesordercheck_v3.datafeed,1)
    
    % Get the case ID
    caseid = char(string(salesordercheck_v3.datafeed(cr,col_df_caseid)));
    
    if strcmp('RS20-IRE-XUP',caseid) == 1
        test = 1;
    end
    
    % Get the position of this case ID in the online report
    row_or_caseid = catchrowindex({caseid},onlinereport,col_or_caseid);
    row_or_caseid = cell2mat(row_or_caseid(2,1));
    
    %     if strcmp(values.y,'20') && strcmp(values.mo,'07') && strcmp(values.d,'07')
    %         % Correct the statuses that were not updated - will only work on the above mentioned day - one time correction
    %         salesordercheck_v3.datafeed(cr,col_df_casestatus) =  onlinereport(row_or_caseid,col_or_casestatus);
    %     end
    
    % Check if there is a shipped date
    temp_shipped = char(string(salesordercheck_v3.datafeed(cr,col_df_shipped)));
    % Check if there is a cancellation date
    temp_cancel = char(string(salesordercheck_v3.datafeed(cr,col_df_casecancelled)));
    if strcmp(temp_shipped,'-') == 1 || strcmp(temp_cancel,'-') == 0
        disp(['Checking (' num2str(cr) '): ' caseid ]);
        
        % Update the status
        salesordercheck_v3.datafeed(cr,col_df_casestatus) =  onlinereport(row_or_caseid,col_or_casestatus);
        
        % Update the cancellation date
        salesordercheck_v3.datafeed(cr,col_df_casecancelled) =  onlinereport(row_or_caseid,col_or_casecancelled);
        
        % Check the production date
        temp = char(string(onlinereport(row_or_caseid,col_or_production)));
        if strcmp(temp,'-') == 0
            salesordercheck_v3.datafeed(cr,col_df_production) =  {temp};
        end
        clear temp
        
        % Check the built date
        temp = char(string(onlinereport(row_or_caseid,col_or_built)));
        if strcmp(temp,'-') == 0
            salesordercheck_v3.datafeed(cr,col_df_built) =  {temp};
        end
        clear temp
        
        % Check the ready to ship date
        temp = char(string(onlinereport(row_or_caseid,col_or_readytoship)));
        if strcmp(temp,'-') == 0
            salesordercheck_v3.datafeed(cr,col_df_readytoship) =  {temp};
        end
        clear temp
        
        % Check the shipped date
        temp = char(string(onlinereport(row_or_caseid,col_or_shipped)));
        if strcmp(temp,'-') == 0
            salesordercheck_v3.datafeed(cr,col_df_shipped) =  {temp};
            
            salesordercheck_v3.datafeed(cr,col_df_shippedyear) =  onlinereport(row_or_caseid,col_or_shippedyear);
            salesordercheck_v3.datafeed(cr,col_df_shippedmonth) =  onlinereport(row_or_caseid,col_or_shippedmonth);
            salesordercheck_v3.datafeed(cr,col_df_shippedday) =  onlinereport(row_or_caseid,col_or_shippedday);
        end
        clear temp
        
    end
end


save([values.salesordercheck_folder '\salesordercheck_v3.mat'],'salesordercheck_v3')

datasets = {'go4d','superfeet'};
nrofdatasets = size(datasets,2);

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%                       IMPORTANT                           %
%                                                           %
% The superfeet orders are still included in the G04D file, %
% but they are removed in the RS Scan portal                %
%                                                           %
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


% Split up the GO4D and superfeet file
for cds = 1:nrofdatasets
    cds_name = char(datasets(1,cds));
    eval(['OR_' cds_name ' = salesordercheck_v3.datafeed;']);
    
    disp(['Preparing the ' cds_name ' file.']);
    

    
    if strcmp(cds_name,'superfeet') == 1
        OR_superfeet(2:12035,:) = [];
        nrofcolsor = size(OR_superfeet,2);
        headers = {'Surgeon','PatientBirthDate','CustomerFirstName','CustomerSurname'};
        % Delete the unwanted columns
        deleted = 0;
        for x = 1:nrofcolsor
            ch = char(OR_superfeet(1,x-deleted));
            if isempty(find(contains(headers,ch))) == 0
                % Delete column
                OR_superfeet(:,x-deleted) = [];
                deleted = deleted + 1;
            else
                % Do nothing
            end
        end
    end
    
    
    filename = [values.onlinereportfolder '20' values.y values.mo values.d '_' values.h values.mi values.s '_' cds_name '.csv'];    %must end in csv
    writetable(cell2table(eval(['OR_' cds_name])), filename, 'writevariablenames', false, 'quotestrings', true);
    if strcmp(values.version,'liveversion') == 1
        % Activate the python script to send csv file to the footscan portal.
        command = ['C:\Users\mathlab\AppData\Local\Programs\Python\Python310\python.exe ' values.currentversion 'Soft\update-rsprint-orders.py ' cds_name ' ' filename];
        system(command);
    else
        command = ['C:\Users\mathlab\AppData\Local\Programs\Python\Python310\python.exe ' values.currentversion 'Soft\update-rsprint-orders-test.py ' cds_name ' ' filename];
        system(command);
    end
    
end


end