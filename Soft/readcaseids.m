function readcaseids()

load Temp\values.mat values

caseids_production = cell(0,1);

if (values.FinRepLabels == 1 && values.UpdateOnlineReport == 0) || values.GrabSalesOrdersUser == 1 
    % Read the file with the case IDs
    if values.NewCaseIDsPresent == 0 
        files = dir('Input\*MTLS*.xlsx');
        directory.caseids = [files(1).folder '\' files(1).name];
        [brol2,caseids_production,brol1] = xlsread(directory.caseids,1);
        clear files        
    end    
    % Read the manual cases
    if values.NewCaseIDsPresent == 1 || values.GrabSalesOrdersUser == 1
        [caseids_production] = readmanualcaseids(caseids_production,'NewCaseID',1,'',1); % Read from input field, ignore rs-numbers
    end
    caseids_production(ismember(caseids_production,'RS19-PHI-ITS')) = [];
    % Sort alphabetically
    caseids_production = sortrows(caseids_production,1);
    xlswrite('Temp\caseids.xlsx',caseids_production)
    save Temp\caseids_production.mat caseids_production
    
end

% Read the cases to (1) prepare the delivery note, (2) to prepare the shipments or (3) to open the logfile(s)
if ((values.ManDelNote == 1 || values.SortOutDelNoteWSxml == 1 || values.FDAlabels == 1 || values.OpenLogFile == 1) && values.NewCaseIDsPresent == 1) ...
        || (values.DailySnapshot == 1 && isfile([pwd '\Temp\SnapshotDone.txt']) == 1)
    [caseids_production] = readmanualcaseids(caseids_production,'NewCaseID',1,'',0); % Read from input field, do not ignore rs-numbers
    caseids_production(ismember(caseids_production,'RS19-PHI-ITS')) = [];
    % Sort alphabetically
    caseids_production = sortrows(caseids_production,1);
    xlswrite('Temp\caseids.xlsx',caseids_production)
    save Temp\caseids_production.mat caseids_production
end

% Read the case for the daily snapshot of production.
if values.DailySnapshot == 1 && isfile([pwd '\Temp\cp.mat']) == 1 && isfile([pwd '\Temp\SnapshotDone.txt']) == 0
    load Temp\cp.mat cp
    %cp_text = char(values.snapshotprocedures_text(1,cp));
    cp_ids = eval(['values.' char(values.snapshotprocedures(1,cp))]);
    [caseids_production] = readmanualcaseids(caseids_production,'',0,cp_ids,0);
    xlswrite('Temp\caseids.xlsx',caseids_production)
    save Temp\caseids_production.mat caseids_production
end

% Check for duplicate case ID's and abort if necessary.
if values.DailySnapshot == 0
    checkforduplicates(caseids_production);
end


end