function GrabSalesOrdersUPS()

load Temp\values.mat values
load Temp\caseids_production.mat caseids_production

nog aanpassen zodat het overeenkomt met de nav ship module

originfolder = [values.backupfolder 'NavisionSalesOrders\All'];
selectionfolder = [pwd '\Output\Navision'];
fetchedfolder = [values.backupfolder 'NavisionSalesOrders\Fetched'];

nrofcaseids = size(caseids_production,1);
success = 0;

for ccid = 1:nrofcaseids
    cid = char(caseids(ccid,1));
    % Check if a file is present
    try
        copyfile([originfolder '\*' cid '.xml'],selectionfolder);
        copyfile([originfolder '\*' cid '.xml'],fetchedfolder);
        success = success + 1;
    catch
        disp(['Grabbing xml sales order xml file did not work for ' cid]);
    end
end
disp(['Success = ' num2str(success) '/' num2str(nrofcaseids)]);

end