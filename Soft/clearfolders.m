function clearfolders

folders = {'Output\Labels','Output\DeliveryNotes','Output\FinishingReports','Output\FinishingReports\BE','Output\FinishingReports\US',...
    'Output\Shipments','Output\Invoices','Temp','Output\WorldShip','C:\Users\mathlab\Downloads','Output\Navision','Output\CaseIDs','Output\Production'};

nroffolders = size(folders,2);
for x = 1:nroffolders
    clc
    whichfolder = char(folders(1,x));
    dinfo = dir(fullfile(whichfolder,'*.*'));
    for K = 1 : length(dinfo)
        thisfile = fullfile(whichfolder, dinfo(K).name);
        delete(thisfile);
    end
end
clc
   
% clc
% whichfolder = 'Output\Labels';
% dinfo = dir(fullfile(whichfolder,'*.*'));
% for K = 1 : length(dinfo)
%     thisfile = fullfile(whichfolder, dinfo(K).name);
%     delete(thisfile);
% end
% clc
% whichfolder = 'Output\DeliveryNotes';
% dinfo = dir(fullfile(whichfolder,'*.*'));
% for K = 1 : length(dinfo)
%     thisfile = fullfile(whichfolder, dinfo(K).name);
%     delete(thisfile);
% end
% clc
% whichfolder = 'Output\FinishingReports';
% dinfo = dir(fullfile(whichfolder,'*.*'));
% for K = 1 : length(dinfo)
%     thisfile = fullfile(whichfolder, dinfo(K).name);
%     delete(thisfile);
% end
% clc
% whichfolder = 'Output\FinishingReports\BE';
% dinfo = dir(fullfile(whichfolder,'*.*'));
% for K = 1 : length(dinfo)
%     thisfile = fullfile(whichfolder, dinfo(K).name);
%     delete(thisfile);
% end
% clc
% whichfolder = 'Output\FinishingReports\US';
% dinfo = dir(fullfile(whichfolder,'*.*'));
% for K = 1 : length(dinfo)
%     thisfile = fullfile(whichfolder, dinfo(K).name);
%     delete(thisfile);
% end
% clc
% whichfolder = 'Output\Shipments';
% dinfo = dir(fullfile(whichfolder,'*.*'));
% for K = 1 : length(dinfo)
%     thisfile = fullfile(whichfolder, dinfo(K).name);
%     delete(thisfile);
% end
% clc
% whichfolder = 'Output\Invoices';
% dinfo = dir(fullfile(whichfolder,'*.*'));
% for K = 1 : length(dinfo)
%     thisfile = fullfile(whichfolder, dinfo(K).name);
%     delete(thisfile);
% end
% clc
% whichfolder = 'Temp';
% dinfo = dir(fullfile(whichfolder,'*.*'));
% for K = 1 : length(dinfo)
%     thisfile = fullfile(whichfolder, dinfo(K).name);
%     delete(thisfile);
% end
% clc
% whichfolder = 'Output\WorldShip';
% dinfo = dir(fullfile(whichfolder,'*.*'));
% for K = 1 : length(dinfo)
%     thisfile = fullfile(whichfolder, dinfo(K).name);
%     delete(thisfile);
% end
% clc
% whichfolder = 'C:\Users\mathlab\Downloads';
% dinfo = dir(fullfile(whichfolder,'*.*'));
% for K = 1 : length(dinfo)
%     thisfile = fullfile(whichfolder, dinfo(K).name);
%     delete(thisfile);
% end
% clc
% whichfolder = 'Output\Navision';
% dinfo = dir(fullfile(whichfolder,'*.*'));
% for K = 1 : length(dinfo)
%     thisfile = fullfile(whichfolder, dinfo(K).name);
%     delete(thisfile);
% end
% clc
% whichfolder = 'Output\CaseIDs';
% dinfo = dir(fullfile(whichfolder,'*.*'));
% for K = 1 : length(dinfo)
%     thisfile = fullfile(whichfolder, dinfo(K).name);
%     delete(thisfile);
% end
% clc

end