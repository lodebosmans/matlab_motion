function openlogfile()

disp('Opening logfiles - please wait');

load Temp\values.mat values
load Temp\caseids_production.mat caseids_production

nrofcases = size(caseids_production,1);
for ccid = 1:nrofcases
    ccid_str = char(caseids_production(ccid,1));
    filename = [values.backupfolder 'LogFiles\Cases\' ccid_str '.txt'];
    if isfile(filename) == 1
        system(filename);
        %system('Temp\cl_batch.bat');
    else
        disp(['A logfile for ' ccid_str ' does not exist.']);
    end
end

end