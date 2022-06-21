function sortshipmentexceltofolder(info)

% % Get the cases from the file.
% [num,text,raw] = xlsread(info.newfile);
% nrofcases = size(text,1);

% % If they need to be logged in the UPS Excel, get the case IDs
% if strcmp(info.cstn,'ignore') == 0
%     % Write them in the shipments variable.
%     eval(['shipments.shipment_' info.cstn '.deliverynumber = info.cstn;']);
%     % Add the case IDs to the shipment
%     eval(['shipments.shipment_' info.cstn '.caseids = text;']);
% end

% Check if all necessary folders exist
if isfolder(info.destinationfolder_sorted) == 0
    mkdir(info.destinationfolder_sorted);
else
    % Folder exists, do nothing.
end

% Copy to the all folder
copyfile(info.newfile,[info.destinationfolder_sorted '\' info.filename])

% Cut and move to the sorted folder
movefile(info.newfile,[info.destinationfolder_all '\' info.filename])

end