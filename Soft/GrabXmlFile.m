clear all 
close all
clc

idfile = 'caseids.xlsx';
foldername = 'O:\Administration\BackupFiles\Navision\SalesOrders\All';
destinationfolder = 'O:\Administration\BackupFiles\Navision\SalesOrders\Fetched';

%%

disp('Start');

[data.num,data.txt,caseids] = xlsread([destinationfolder '\' idfile],1);

% listing = dir(foldername);
nrofcaseids = size(caseids,1);
Success = 0;
failedcases = cell(0);

for ccid = 1:nrofcaseids
    cid = char(caseids(ccid,1));
    % Check if a file is present
    try
        copyfile([foldername '\*' cid '.xml'],destinationfolder)
        Success = Success +1;
        disp(['Grabbed ' cid ' (' num2str(ccid) '/' num2str(nrofcaseids) ')']);
    catch
        disp(['Did not work for ' cid])
        nextfail = size(failedcases,1) + 1;
        failedcases(nextfail,1) = {cid};        
    end
end
disp(['Success = ' num2str(Success) '/' num2str(nrofcaseids)]);

if size(failedcases,1) > 0
    disp('Failed cases:');
    for x = 1:size(failedcases,1)
        disp(char(failedcases(x,1)));        
    end
else
    disp('No failed cases.');
end
disp('Finish');