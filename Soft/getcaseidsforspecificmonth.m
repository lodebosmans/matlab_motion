function [caseidsmonth] = getcaseidsforspecificmonth()

load Temp\values.mat values
load Temp\UPSfile_shipment.mat UPSfile_shipment

caseidsmonth = cell(0);

upsfile = UPSfile_shipment;
col_SerialNumbers = catchcolumnindex({'SerialNumbers'},upsfile,1);
col_SerialNumbers = cell2mat(col_SerialNumbers(2,1));
col_Service = catchcolumnindex({'Service'},upsfile,1);
col_Service = cell2mat(col_Service(2,1));

% Get the columnnumbers
col_Year = catchcolumnindex({'Year'},upsfile,1);
col_Year = cell2mat(col_Year(2,1));
col_Month = catchcolumnindex({'Month'},upsfile,1);
col_Month = cell2mat(col_Month(2,1));
col_Day = catchcolumnindex({'Day'},upsfile,1);
col_Day = cell2mat(col_Day(2,1));
col_ShipmentNumber = catchcolumnindex({'ShipmentNumber'},upsfile,1);
col_ShipmentNumber = cell2mat(col_ShipmentNumber(2,1));
col_Trackingnumber = catchcolumnindex({'Trackingnumber'},upsfile,1);
col_Trackingnumber = cell2mat(col_Trackingnumber(2,1));
col_Reference = catchcolumnindex({'Reference'},upsfile,1);
col_Reference = cell2mat(col_Reference(2,1));

% Go over all the rows in the shipment excel
% If the date matches, continue
nrofrows = size(upsfile,1);
% Skip the first row with the headers
for cr = 2:nrofrows
    % If the date of shipment matches the selected date's and the shipmentnumber is not empty, continue.
    if cell2mat(upsfile(cr,col_Year)) == str2double(values.RequestedYear) ...
            && cell2mat(upsfile(cr,col_Month)) == str2double(values.RequestedMonth) ...
            && strcmp(cell2mat(upsfile(cr,col_ShipmentNumber)),'-') == 0
        
        % Get the case IDs from that shipment
        currentshipment = char(upsfile(cr,col_SerialNumbers));        
        [caseids] = readmanualcaseids({},'',0,currentshipment,1);  
        nrofcasesinshipment = size(caseids,1);
        
        % Add to the list of case IDs
        crows = size(caseidsmonth,1);
        startposition = crows + 1;
        endposition = crows + nrofcasesinshipment;
        caseidsmonth(startposition:endposition,1) = caseids;
    end
end

end