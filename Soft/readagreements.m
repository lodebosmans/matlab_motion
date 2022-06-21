function readagreements()

load Temp\values.mat values

disp('Reading customer agreement details - please wait');
[data.num,data.txt,UPSfile_customagreements] = xlsread([values.backupfolder 'Navision\Agreements\' values.agreementsfile],1);
% Filter out the NaNs
UPSfile_customagreements(cellfun(@(x) isnumeric(x) && isnan(x), UPSfile_customagreements)) = {'-'};
save Temp\UPSfile_customagreements.mat UPSfile_customagreements;

% Backup agreements
copyfile([values.backupfolder 'Navision\Agreements\' values.agreementsfile],'Temp\Agreements.xlsx');
% Copy for Polle
copyfile([values.backupfolder 'Navision\Agreements\' values.agreementsfile],[values.pollespowerbifolder values.agreementsfile]);

end