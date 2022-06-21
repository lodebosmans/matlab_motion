function grabxmlfilesuser()

load Temp\values.mat values
load Temp\caseids_production.mat caseids_production

disp('Grabbing some sales order xml files - please wait');

originfolder = [values.backupfolder 'Navision\SalesOrders\All'];
selectionfolder = [pwd '\Output\Navision'];

nrofcaseids = size(caseids_production,1);
success = 0;
failure = 0;
failedcases = cell(0);

for ccid = 1:nrofcaseids
    cid = char(caseids_production(ccid,1));
    % Check if a file is present
    try
        copyfile([originfolder '\*' cid '.xml'],selectionfolder);
        success = success + 1;
        disp(['Grabbed ' cid ' sales order xml ' num2str(ccid) '/' num2str(nrofcaseids)])
    catch
        disp(['Grabbing xml sales order xml file did not work for ' cid]);
        failure = failure + 1;
        failedcases(failure,1) = {cid};
    end
end
disp(' ')
disp(['Success = ' num2str(success) '/' num2str(nrofcaseids)]);
disp(' ')

if failure > 0
    disp('Failed cases:');
    for x = 1:failure
        disp(char(failedcases(x,1)));
    end
end

end