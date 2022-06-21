function [shipments] = getshippedcaseids(info,shipments)

% Get the cases from the file.
[num,text,raw] = xlsread(info.newfile);


% If they need to be logged in the UPS Excel, get the case IDs
if strcmp(info.cstn,'ignore') == 0
    % Write them in the shipments variable.
    eval(['shipments.shipment_' info.cstn '.deliverynumber = info.cstn;']);
    % Add the case IDs to the shipment
    eval(['shipments.shipment_' info.cstn '.caseids = text;']);
end

end