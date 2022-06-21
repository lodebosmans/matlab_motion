function [info] = writedeliverynotetofile(info,condition,customers,onlinereport,salesordercheck_v3,UPSfile_shipment,cr)

load Temp\values.mat values

% Columns SOC
col_caseid = catchcolumnindex({'CaseID'},salesordercheck_v3.headers,1);
col_caseid = cell2mat(col_caseid(2,1));
col_itemnumber = catchcolumnindex({'ItemNumber'},salesordercheck_v3.headers,1);
col_itemnumber = cell2mat(col_itemnumber(2,1));
% Columns onlinereport
col_reference = catchcolumnindex({'ReferenceID'},onlinereport,1);
col_reference = cell2mat(col_reference(2,1));
col_casecode = catchcolumnindex({'CaseCode'},onlinereport,1);
col_casecode = cell2mat(col_casecode(2,1));

% Write to file
% -------------
file = [values.deliverynotetemplatefolder values.deliverynotetemplatefilename]; % This must be full path name  => only adapt the file in the 'Interface' folder, not in the test environment.
% Open Excel Automation server
Excel = actxserver('Excel.Application');
Workbooks = Excel.Workbooks;
% Make Excel visible
Excel.Visible = 1;
% Open Excel file
Workbook = Workbooks.Open(file);
Sheets = Excel.ActiveWorkBook.Sheets;
sheet1 = get(Sheets, 'Item', 1);
invoke(sheet1, 'Activate');
Activesheet = Excel.Activesheet;

% If shipped from US, overwrite the RS PRint address with the MATUS address.
if strcmp(values.CountryOfShipment,'US') == 1
    range = 'A46:A46';
    data = 'Materialise US - 44650 Helm Court, 48170 Plymouth MI, USA - T:17342596445 - support@rsprint.be - VAT 0551855071';
    ActivesheetRange = get(Activesheet,'Range',range);
    set(ActivesheetRange, 'Value', data);
end

% Insert the date
range = 'I4:I4';
if strcmp(condition,'UpsShipment') == 1
    year = num2str(cell2mat(UPSfile_shipment(cr,1)));
    month = num2str(cell2mat(UPSfile_shipment(cr,2)));
    day = num2str(cell2mat(UPSfile_shipment(cr,3)));
    data = [day '/' month '/' year];
elseif strcmp(condition,'ManualShipment') == 1
    data = [values.d '/' values.mo '/' values.y];
end
% Write date to file
ActivesheetRange = get(Activesheet,'Range',range);
set(ActivesheetRange, 'Value', data);

% Insert all other data
col.shipmentlabel = {'ShipmentNumber','CustomerNumber','Weight','Reference'};
col.range = {'D4:D4','G12:G12','H6:H6','G9:G9'};
nrofthingstowrite = size(col.shipmentlabel,2);
for ct = 1:nrofthingstowrite
    % Get the range
    range = char(col.range(1,ct));
    % Get the label
    label = char(col.shipmentlabel(1,ct));
    
    % Get the position of the label in the shipment overview
    cc = catchcolumnindex({label},UPSfile_shipment,1);
    cc = cell2mat(cc(2,1));
    % Fetch the data
    data = UPSfile_shipment(cr,cc);
    if strcmp(label,'ShipmentNumber') == 1
        shipmentnumber = num2str(cell2mat(data));
    end
    % Change to char
    if isnumeric(data)
        data = num2str(data);
    elseif iscell(data)
        data = num2str(cell2mat(data));
    elseif isstring(data)
        data = char(data);
    else
        clear all
    end
    % Check if it is empty. If yes, add dashed line.
    if strcmp(data,'') == 1
        data = '-';
    end
    if strcmp(label,'Insole') == 1
        if strcmp(data(end:end),',')
            data = data(1:length(data)-1);
        end
    end
    % Write to file
    ActivesheetRange = get(Activesheet,'Range',range);
    set(ActivesheetRange, 'Value', data);
    
    %                 % Remove reference if it is a suborder
    %                 if isempty(cell2mat(UPSfile_shipment(cr,col_shipnr)))
    %                     ActivesheetRange = get(Activesheet,'Range','A20:A20');
    %                     set(ActivesheetRange, 'Value', '');
    %                     ActivesheetRange = get(Activesheet,'Range','A21:A21');
    %                     set(ActivesheetRange, 'Value', '');
    %                 end
    
    % Check for special things to do
    if strcmp(label,'CustomerNumber') == 1
        ccnr = data;
        ccnr_ = strrep(ccnr,'-','_');
        if contains(ccnr,'0002') == 0
            
            tempcomp = customers.(ccnr_).DeliveryAddress.DeliveryFacilityName;
            
            if strcmp(customers.(ccnr_).ActionMessage,'-') == 0
                info.specialsamount = info.specialsamount + 1;
                info.specialsthingstodo(info.specialsamount,1) = {data};
                info.specialsthingstodo(info.specialsamount,2) = {tempcomp};
                info.specialsthingstodo(info.specialsamount,3) = {customers.(ccnr_).ActionMessage};
            end
            
            if strcmp(customers.(ccnr_).PrintCommercialInvoice,'Yes') == 1
                info.specialsamount = info.specialsamount + 1;
                info.specialsthingstodo(info.specialsamount,1) = {data};
                info.specialsthingstodo(info.specialsamount,2) = {tempcomp};
                info.specialsthingstodo(info.specialsamount,3) = {'Add the commercial invoice to the shipment and/or email it to the necessary people.'};
            end
            
            % Compare the fixed shipment day with the current date
            if strcmp(customers.(ccnr_).FixedShipmentDay,'-') == 0
                [DayNumber,DayName] = weekday([values.d '-' values.mo '-20' values.y]);
                if DayNumber == 1
                    fixedday = 'Sunday';
                elseif DayNumber == 2
                    fixedday = 'Monday';
                elseif DayNumber == 3
                    fixedday = 'Tuesday';
                elseif DayNumber == 4
                    fixedday = 'Wednesday';
                elseif DayNumber == 5
                    fixedday = 'Thursday';
                elseif DayNumber == 6
                    fixedday = 'Friday';
                elseif DayNumber == 7
                    fixedday = 'Saturday';
                else
                    error('No dayname found.')
                end
                if strcmp(upper(customers.(ccnr_).FixedShipmentDay),upper(fixedday)) == 0
                    info.specialsamount = info.specialsamount + 1;
                    info.specialsthingstodo(info.specialsamount,1) = {data};
                    info.specialsthingstodo(info.specialsamount,2) = {tempcomp};
                    info.specialsthingstodo(info.specialsamount,3) = {['Shipments to ' tempcomp ' are normally planned on ' customers.(ccnr_).FixedShipmentDay '. Please double check if this shipment needs to go out today or not.']};
                end
            end
            
            if strcmp(customers.(ccnr_).NeutralShipmentBox,'Yes') == 1
                info.specialsamount = info.specialsamount + 1;
                info.specialsthingstodo(info.specialsamount,1) = {data};
                info.specialsthingstodo(info.specialsamount,2) = {tempcomp};
                info.specialsthingstodo(info.specialsamount,3) = {'Use neutral shipment box.'};
            end
            
            
            %         if strcmp(data,'C1020') == 1 || strcmp(data,'C1020-001') == 1
            %             % Tili
            %             info.specialsamount = info.specialsamount + 1;
            %             info.specialsthingstodo(info.specialsamount,1) = {data};
            %             info.specialsthingstodo(info.specialsamount,2) = {tempcomp};
            %             info.specialsthingstodo(info.specialsamount,3) = {'Inform support about the shipment.'};
            %         end
            %         if strcmp(data,'C1074') == 1 || strcmp(data,'C1074-001') == 1
            %             % Sinai
            %             info.specialsamount = info.specialsamount + 1;
            %             info.specialsthingstodo(info.specialsamount,1) = {data};
            %             info.specialsthingstodo(info.specialsamount,2) = {tempcomp};
            %             info.specialsthingstodo(info.specialsamount,3) = {'Inform Natascha about the shipment.'};
            %         end
            %         if strcmp(data,'C1015') == 1 || strcmp(data,'C1015-001') == 1 || strcmp(data,'C1015-002') == 1
            %             % Alchemaker
            %             info.specialsamount = info.specialsamount + 1;
            %             info.specialsthingstodo(info.specialsamount,1) = {data};
            %             info.specialsthingstodo(info.specialsamount,2) = {tempcomp};
            %             info.specialsthingstodo(info.specialsamount,3) = {'Add commercial invoice to shipment.'};
            %         end
            %         if strcmp(data,'C1019') == 1 || strcmp(data,'C1019-002') == 1
            %             % RSScan Beijing
            %             info.specialsamount = info.specialsamount + 1;
            %             info.specialsthingstodo(info.specialsamount,1) = {data};
            %             info.specialsthingstodo(info.specialsamount,2) = {tempcomp};
            %             info.specialsthingstodo(info.specialsamount,3) = {'Add commercial invoice to shipment.'};
            %         end
            %         if strcmp(data,'C1142') == 1 || strcmp(data,'C1142-001') == 1
            %             % Benta
            %             info.specialsamount = info.specialsamount + 1;
            %             info.specialsthingstodo(info.specialsamount,1) = {data};
            %             info.specialsthingstodo(info.specialsamount,2) = {tempcomp};
            %             info.specialsthingstodo(info.specialsamount,3) = {'Send commercial invoice to exportlummen@ups.com.'};
            %         end
        end
    end
end

% Get data from the customers variable and add it to the Excel file
% 'InvoiceCompany'
ActivesheetRange = get(Activesheet,'Range','D6:D6');
set(ActivesheetRange, 'Value', customers.(ccnr_).Company);
% 'InvoiceAddress1'
ActivesheetRange = get(Activesheet,'Range','D7:D7');
set(ActivesheetRange, 'Value', customers.(ccnr_).InvoiceAddress.InvoiceAddress1);
% 'InvoiceAddress2'
ActivesheetRange = get(Activesheet,'Range','D8:D8');
set(ActivesheetRange, 'Value', customers.(ccnr_).InvoiceAddress.InvoiceAddress2);
% 'InvoiceAddress3'
ActivesheetRange = get(Activesheet,'Range','D9:D9');
set(ActivesheetRange, 'Value', customers.(ccnr_).InvoiceAddress.InvoiceAddress3);
% 'InvoicePostalCode'
ActivesheetRange = get(Activesheet,'Range','D10:D10');
set(ActivesheetRange, 'Value', customers.(ccnr_).InvoiceAddress.InvoicePostalCode);
% 'InvoiceCity'
ActivesheetRange = get(Activesheet,'Range','D11:D11');
set(ActivesheetRange, 'Value', customers.(ccnr_).InvoiceAddress.InvoiceCity);
% 'InvoiceState'
ActivesheetRange = get(Activesheet,'Range','D12:D12');
set(ActivesheetRange, 'Value', customers.(ccnr_).InvoiceAddress.InvoiceState);
% 'InvoiceCountry'
ActivesheetRange = get(Activesheet,'Range','D13:D13');
set(ActivesheetRange, 'Value', customers.(ccnr_).InvoiceAddress.InvoiceCountry);
% 'DeliveryFacilityName'
ActivesheetRange = get(Activesheet,'Range','A6:A6');
set(ActivesheetRange, 'Value', customers.(ccnr_).DeliveryAddress.DeliveryFacilityName);
% 'DeliveryAddress1'
ActivesheetRange = get(Activesheet,'Range','A7:A7');
set(ActivesheetRange, 'Value', customers.(ccnr_).DeliveryAddress.DeliveryAddress1);
% 'DeliveryAddress2'
ActivesheetRange = get(Activesheet,'Range','A8:A8');
set(ActivesheetRange, 'Value', customers.(ccnr_).DeliveryAddress.DeliveryAddress2);
% 'DeliveryAddress3'
ActivesheetRange = get(Activesheet,'Range','A9:A9');
set(ActivesheetRange, 'Value', customers.(ccnr_).DeliveryAddress.DeliveryAddress3);
% 'DeliveryPostalCode'
ActivesheetRange = get(Activesheet,'Range','A10:A10');
set(ActivesheetRange, 'Value', customers.(ccnr_).DeliveryAddress.DeliveryPostalCode);
% 'DeliveryCity'
ActivesheetRange = get(Activesheet,'Range','A11:A11');
set(ActivesheetRange, 'Value', customers.(ccnr_).DeliveryAddress.DeliveryCity);
% 'DeliveryState'
ActivesheetRange = get(Activesheet,'Range','A12:A12');
set(ActivesheetRange, 'Value', customers.(ccnr_).DeliveryAddress.DeliveryState);
% 'DeliveryCountry
ActivesheetRange = get(Activesheet,'Range','A13:A13');
set(ActivesheetRange, 'Value', customers.(ccnr_).DeliveryAddress.DeliveryCountry);

% If it is not a manual delivery, remove the signature area
if strcmp(condition,'ManualShipment') == 0
    ActivesheetRange = get(Activesheet,'Range','D40:D44');
    set(ActivesheetRange, 'Value', '');
end

% Read the amount of cases
sn_col = catchcolumnindex({'SerialNumbers'},UPSfile_shipment,1);
sn_col = cell2mat(sn_col(2,1));
nonsn_col = catchcolumnindex({'NonSerialNumbers'},UPSfile_shipment,1);
nonsn_col = cell2mat(nonsn_col(2,1));
[output] = readmanualcaseids({},'',0,char(UPSfile_shipment(cr,sn_col)),0);
nonsnrow = 0;
[output_nsn] = UPSfile_shipment(cr,nonsn_col);
nrofcases = size(output,1);
% Add the itemnumber and reference
if nrofcases > 0
    for currentlineindeliverynote = 1:nrofcases
        % Get the case ID
        temp_ccid = output(currentlineindeliverynote,1);
        % Only proceed if it is a regular case ID - RS-numbers will be igonored
        if (contains(char(temp_ccid),'RS1') || contains(char(temp_ccid),'RS2')) == 1
            % Find it in the SOC
            SOC_current_ID = catchrowindex(temp_ccid,salesordercheck_v3.data,1);
            SOC_current_ID = cell2mat(SOC_current_ID(2,1));
            % Get the itemnumber from the SOC
            output(currentlineindeliverynote,2) = salesordercheck_v3.data(SOC_current_ID,col_itemnumber);
            % Get the reference ID from the online report
            OR_current_row = catchrowindex(temp_ccid,onlinereport,col_casecode);
            OR_current_row = cell2mat(OR_current_row(2,1));
            output(currentlineindeliverynote,3) = onlinereport(OR_current_row,col_reference);
        else
            output(currentlineindeliverynote,2) = {'-'};
            output(currentlineindeliverynote,3) = {'-'};
        end
    end
else
    nrofcases = 1;
    output(1,1) = output_nsn;
    %     output(1,2) = {'-'};
    %     output(1,3) = {'-'};
    nonsnrow = 1;
end

% Excel template parameters
page.rows = 46; % Number of row on the entire page of the template
page.columns = 'I'; % Index of the total amount of colums in the entire page the template
page.startindex = 15; % Relative index of the startposition where the row where the invoice data will be pasted
page.endindex = 39; % Relative index of the endposition where the row where the invoice data will be pasted
page.casesperpage = page.endindex - page.startindex + 1; % Number of cases per page in the template
page.indexcol = 'A'; % Column where the numbering will be added
template = ['A' num2str(1) ':' page.columns num2str(page.rows)];
nrofpages = ceil(nrofcases/25);

for newpage = 1:nrofpages
    % Find the start and end position to paste the template
    startrow = 1 + (page.rows*(newpage-1));
    endrow = 46 + (page.rows*(newpage-1));
    if newpage == 1
        % Do nothing
    else
        % Get the template
        sheet1.Range(template).Copy;
        % Define the destination location and paste the template
        destination = ['A' num2str(startrow) ':' page.columns num2str(endrow)];
        sheet1.Paste(sheet1.Range(destination));
    end
    % Insert the page number
    ActivesheetRange = get(Activesheet,'Range',[page.indexcol num2str(startrow) ':' page.indexcol num2str(startrow)]);
    set(ActivesheetRange, 'Value', ['Page ' num2str(newpage) '/' num2str(nrofpages)]);
    % Center the bottom line with the company information
    mergelocation = [page.indexcol num2str(endrow) ':' page.columns num2str(endrow)];
    sheet1.Range(mergelocation).Merge
    %                 ActivesheetRange = get(Activesheet,'Range',mergelocation);
    %                 set(ActivesheetRange, 'HorizontalAlignment', xlCenter);
    cells = Excel.ActiveSheet.Range(mergelocation);
    cells.Select;
    cells.MergeCells = 1;
    set(cells,'HorizontalAlignment',3);
    
end

for newpage = 1:nrofpages
    
    % Insert the data
    if nrofpages == 1
        startrow = page.startindex + (page.rows*(newpage-1));
        endrow = startrow + nrofcases - (newpage-1)*page.casesperpage  - 1;
        startdata = (newpage-1)* page.casesperpage + 1;
        enddata = startdata + nrofcases - (newpage-1)*page.casesperpage - 1;
    elseif newpage < nrofpages
        startrow = page.startindex + (page.rows*(newpage-1));
        endrow = page.endindex + (page.rows*(newpage-1));
        startdata = (newpage-1)* page.casesperpage + 1;
        enddata = startdata + page.casesperpage - 1;
    elseif newpage == nrofpages
        startrow = page.startindex + (page.rows*(newpage-1));
        endrow = startrow + nrofcases - (newpage-1)*page.casesperpage  - 1;
%         endrow = startrow + nrofcases - floor(nrofcases/page.casesperpage)*page.casesperpage - 1;
        startdata = (newpage-1)* page.casesperpage + 1;
        enddata = startdata + nrofcases - (newpage-1)*page.casesperpage - 1;
    end
    
    for x = 1:endrow-startrow+1
        % Insert the case ID
        ActivesheetRange = get(Activesheet,'Range',['B' num2str(startrow + x - 1) ':B' num2str(startrow + x - 1)]);
        set(ActivesheetRange, 'Value', output(startdata + x - 1,1));
        if nonsnrow == 0
            % Insert the itemnumber
            ActivesheetRange = get(Activesheet,'Range',['D' num2str(startrow + x - 1) ':D' num2str(startrow + x - 1)]);
            set(ActivesheetRange, 'Value', output(startdata + x - 1,2));
            % Insert the reference
            ActivesheetRange = get(Activesheet,'Range',['G' num2str(startrow + x - 1) ':G' num2str(startrow + x - 1)]);
            set(ActivesheetRange, 'Value', output(startdata + x - 1,3));
        else
            % Remove the useless headers
            temp3 = ['B' num2str(startrow-1) ':I' num2str(startrow-1)];
            sheet1.Range(temp3).Value = '';
            temp3 = ['B' num2str(startrow-1) ':B' num2str(startrow-1)];
            %             sheet1.Range(temp3).Merge
            sheet1.Range(temp3).Value = 'Description';
            % Merge the first row to make space for the description of the content
            temp3 = ['B' num2str(startrow) ':I' num2str(startrow)];
            sheet1.Range(temp3).Merge
        end
    end
    
    % Insert the numbers
    temp = startdata:enddata;
    temp = temp';
    ActivesheetRange = get(Activesheet,'Range',[page.indexcol num2str(startrow) ':' page.indexcol num2str(endrow)]);
    set(ActivesheetRange, 'Value', temp);
end

cc = catchcolumnindex({'ContentDescription'},UPSfile_shipment,1);
cc = cell2mat(cc(2,1));
%             cc2 = catchcolumnindex({'CustomerNumber'},UPSfile_shipment,1);
%             cc2 = cell2mat(cc2(2,1));
if strcmp(UPSfile_shipment(cr,cc),'Pressure plate') == 1
    % Pressure plate
    info.specialsamount = info.specialsamount + 1;
    info.specialsthingstodo(info.specialsamount,1) = {ccnr};
    info.specialsthingstodo(info.specialsamount,2) = {tempcomp};
    info.specialsthingstodo(info.specialsamount,3) = {'Add conformity declaration to shipment with the footscan plate.'};
end

% % Create the output filename
% outputfile = [pwd '\Output\DeliveryNotes\' values.y values.mo values.d '_' values.h values.mi values.s '_DeliveryNote_' customers.(ccnr_).Company];

% Create the output filename
outputfile = [pwd '\Output\DeliveryNotes\' values.y values.mo values.d '_' values.h values.mi values.s '_DeliveryNote_' num2str(cr) '_' shipmentnumber '_' customers.(ccnr_).Company];
% Check if the outputfile already exists or not
while isfile([outputfile '.xlsx']) == 1
    outputfile = [outputfile '_1' ];
end

% Send to printer
if values.MenuSelectPrinter > 1 && values.ManDelNote == 1
    Excel.ActiveWorkbook.PrintOut(1,nrofpages,values.NrOfCopies,'False',values.MenuSelectPrinter_str);
end
% Save as excel
invoke(Workbook,'SaveAs',outputfile);
% Save as pdf
sheet1.ExportAsFixedFormat('xlTypePDF',outputfile);
% Close Excel and clean up
invoke(Excel,'Quit');
delete(Excel);
clear Excel;

end
