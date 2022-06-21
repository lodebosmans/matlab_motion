function createdeliverynoteups()

load Temp\values.mat values
load Temp\UPSfile_shipment.mat UPSfile_shipment
load Temp\customers.mat customers

clc
disp('Creating delivery notes - please wait')

[nrofrows,nrofcols] = size(UPSfile_shipment); %#ok<ASGLU>
shipmentcounter = 0;
%UPSfile_shipment_headers = UPSfile_shipment(1,:);
%UPSfile_shipment_headers(2,:) = {zeros};  

specialsthingstodo = cell(0);
specialsamount = size(specialsthingstodo,1);

% Get the column numbers.
col_shipped = catchcolumnindex({'PromotedOMS'},UPSfile_shipment,1);
col_shipped = cell2mat(col_shipped(2,1));
col_service = catchcolumnindex({'Service'},UPSfile_shipment,1);
col_service = cell2mat(col_service(2,1));
col_shipnr = catchcolumnindex({'ShipmentNumber'},UPSfile_shipment,1);
col_shipnr = cell2mat(col_shipnr(2,1));
col_ShippedFrom = catchcolumnindex({'ShippedFrom'},UPSfile_shipment,1);
col_ShippedFrom = cell2mat(col_ShippedFrom(2,1));

for cr = 2:nrofrows
    if cr == 20
        test = 1;
    end
    % Check is the shipment in the row is already shipped. If not, fetch all necessary data
    %if isempty(cell2mat(UPSfile_shipment(cr,col_shipped))) == 1 && strcmp(UPSfile_shipment(cr,col_service),'UPS') == 1
    if strcmp(cell2mat(UPSfile_shipment(cr,col_shipped)),'-') == 1 && strcmp(UPSfile_shipment(cr,col_service),'UPS') == 1 && strcmp(char(UPSfile_shipment(cr,col_ShippedFrom)),values.CountryOfShipment) == 1
        shipmentcounter = shipmentcounter + 1;        
        
        % Write to file
        % -------------
        file = [values.baseversion 'Input\Templates\DeliveryNoteUpsShipment_v2.xlsx']; % This must be full path name  => only adapt the file in the 'Interface' folder, not in the test environment.
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
            range = 'A1:A1';
            data = 'RS Print (Materialise US)';
            ActivesheetRange = get(Activesheet,'Range',range);
            set(ActivesheetRange, 'Value', data);
            range = 'A2:A2';
            data = '44650 Helm Court';
            ActivesheetRange = get(Activesheet,'Range',range);
            set(ActivesheetRange, 'Value', data);
            range = 'A3:A3';
            data = '48170 Plymouth';
            ActivesheetRange = get(Activesheet,'Range',range);
            set(ActivesheetRange, 'Value', data);
            range = 'A4:A4';
            data = 'Michigan';
            ActivesheetRange = get(Activesheet,'Range',range);
            set(ActivesheetRange, 'Value', data);
            range = 'A5:A5';
            data = 'United States';
            ActivesheetRange = get(Activesheet,'Range',range);
            set(ActivesheetRange, 'Value', data);
            range = 'A6:A6';
            data = 'Tel.: 17342596445';
            ActivesheetRange = get(Activesheet,'Range',range);
            set(ActivesheetRange, 'Value', data);
        end
        
        % Insert the date
        range = 'I8:I8';
        year = num2str(cell2mat(UPSfile_shipment(cr,1)));
        month = num2str(cell2mat(UPSfile_shipment(cr,2)));
        day = num2str(cell2mat(UPSfile_shipment(cr,3)));
        data = [day '/' month '/' year];
        % Write to file
        ActivesheetRange = get(Activesheet,'Range',range);
        set(ActivesheetRange, 'Value', data);
        
        % Insert all other data
        col.shipmentlabel = {'ShipmentNumber','CustomerNumber','Weight','Insole',...
                             'SerialNumbers','Reference','InvoiceCompany',...
                             'InvoiceAddress1','InvoiceAddress2','InvoiceAddress3','InvoicePostalCode',...
                             'InvoiceCity','InvoiceState','InvoiceCountry','DeliveryFacilityName',...
                             'DeliveryAddress1','DeliveryAddress2','DeliveryAddress3','DeliveryPostalCode','DeliveryCity',...
                             'DeliveryState','DeliveryCountry'};
        col.range = {'D8:D8','G21:G21','G11:G11','A24:A24',...
                     'C25:C26','A21:A21','D11:D11',...
                     'D12:D12','D13:D13','D14:D14','D15:D15',...
                     'D16:D16','D17:D17','D18:D18','A11:A11',...
                     'A12:A12','A13:A13','A14:A14','A15:A15','A16:A16',...
                     'A17:A17','A18:A18'};
                
        cc = catchcolumnindex({'DeliveryFacilityName'},UPSfile_shipment,1);
        cc = cell2mat(cc(2,1));
        tempcomp = char(UPSfile_shipment(cr,cc));
        
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
            if strcmp(label,'InvoiceCompany') == 1
                companyname = data;
            end
            if strcmp(label,'Insole') == 1
                if strcmp(data(end:end),',')
                    data = data(1:length(data)-1);
                end
            end
            % Write to file
            ActivesheetRange = get(Activesheet,'Range',range);
            set(ActivesheetRange, 'Value', data);
            % Remove reference if it is a suborder
            if isempty(cell2mat(UPSfile_shipment(cr,col_shipnr)))
                ActivesheetRange = get(Activesheet,'Range','A20:A20');
                set(ActivesheetRange, 'Value', '');
                ActivesheetRange = get(Activesheet,'Range','A21:A21');
                set(ActivesheetRange, 'Value', '');
            end
            
            % Check for special things to do            
            if ct == 2
                if strcmp(data,'C1020') == 1 || strcmp(data,'C1020-001') == 1
                    % Tili
                    specialsamount = specialsamount + 1;
                    specialsthingstodo(specialsamount,1) = {data};
                    specialsthingstodo(specialsamount,2) = {tempcomp};
                    specialsthingstodo(specialsamount,3) = {'Inform support about the shipment.'};
                end
                if strcmp(data,'C1074') == 1 || strcmp(data,'C1074-001') == 1
                    % Sinai
                    specialsamount = specialsamount + 1;
                    specialsthingstodo(specialsamount,1) = {data};
                    specialsthingstodo(specialsamount,2) = {tempcomp};
                    specialsthingstodo(specialsamount,3) = {'Inform Natascha about the shipment.'};
                end
                if strcmp(data,'C1015') == 1 || strcmp(data,'C1015-001') == 1
                    % Alchemaker
                    specialsamount = specialsamount + 1;
                    specialsthingstodo(specialsamount,1) = {data};
                    specialsthingstodo(specialsamount,2) = {tempcomp};
                    specialsthingstodo(specialsamount,3) = {'Add commercial invoice to shipment.'};                    
                end
                if strcmp(data,'C1019') == 1 || strcmp(data,'C1019-002') == 1
                    % RSScan Beijing
                    specialsamount = specialsamount + 1;
                    specialsthingstodo(specialsamount,1) = {data};
                    specialsthingstodo(specialsamount,2) = {tempcomp};
                    specialsthingstodo(specialsamount,3) = {'Add commercial invoice to shipment.'};                    
                end
                if strcmp(data,'C1142') == 1 || strcmp(data,'C1142-001') == 1
                    % Benta
                    specialsamount = specialsamount + 1;
                    specialsthingstodo(specialsamount,1) = {data};
                    specialsthingstodo(specialsamount,2) = {tempcomp};
                    specialsthingstodo(specialsamount,3) = {'Send commercial invoice to exportlummen@ups.com.'};
                end
            end

        end  
        
        cc = catchcolumnindex({'ContentDescription'},UPSfile_shipment,1);
        cc = cell2mat(cc(2,1));             
        cc2 = catchcolumnindex({'CustomerNumber'},UPSfile_shipment,1);
        cc2 = cell2mat(cc2(2,1));  
        if strcmp(UPSfile_shipment(cr,cc),'Pressure plate') == 1
            % Pressure plate
            specialsamount = specialsamount + 1;
            specialsthingstodo(specialsamount,1) = {char(UPSfile_shipment(cr,cc2))};
            specialsthingstodo(specialsamount,2) = {tempcomp};
            specialsthingstodo(specialsamount,3) = {'Add conformity declaration to shipment with the footscan plate'};
        end
        
        % Create the output filename
        outputfile = [pwd '\Output\DeliveryNotes\' values.y values.mo values.d '_' values.h values.mi values.s '_DeliveryNote_' companyname];
        % Check if the outputfile already exists or not
        if isfile([outputfile '.xlsx']) == 1         
            outputfile = [outputfile '_1' ];
        end
        
        % Send to printer
        if values.MenuSelectPrinter > 1
            Excel.ActiveWorkbook.PrintOut(1,1,1,'False',values.MenuSelectPrinter_str);
        end        
        % Save as excel
        invoke(Workbook,'SaveAs',outputfile)
        % Save as pdf
        sheet1.ExportAsFixedFormat('xlTypePDF',outputfile);
        % Close Excel and clean up
        invoke(Excel,'Quit');
        delete(Excel);
        clear Excel;
    end
end

% Print the special things to do
if specialsamount > 0    
    Excel = actxserver('Excel.Application');
    % Set it to visible
    set(Excel,'Visible',1);
    % Add a Workbook
    Workbooks = Excel.Workbooks;
    Workbook = invoke(Workbooks, 'Add');
    % Get a handle to Sheets and select Sheet 1
    Sheets = Excel.ActiveWorkBook.Sheets;
    Sheet1 = get(Sheets, 'Item', 1);
    Sheet1.Activate;
    eActivesheetRange = get(Excel.Activesheet,'Range',['A1:C' num2str(specialsamount)]);
    eActivesheetRange.Value = specialsthingstodo;
    
    Sheet1.Columns.Item(1).columnWidth = 10;
    Sheet1.Columns.Item(2).columnWidth = 20;
    Sheet1.Columns.Item(3).columnWidth = 50;
    
    % And set it to text wrap
    eActivesheetRange = get(Excel.Activesheet,'Range','A:C');
    eActivesheetRange.WrapText = true;
    % Align everything vertically upwards
    eActivesheetRange = get(Excel.Activesheet,'Range','A:C');
    eActivesheetRange.VerticalAlignment = 1;
    
    % Create the output filename
    outputfile = [pwd '\Output\DeliveryNotes\' values.y values.mo values.d '_' values.h values.mi values.s '_OverviewSpecialThings'];

    % Send to printer
    if values.MenuSelectPrinter > 1
        Excel.ActiveWorkbook.PrintOut(1,1,1,'False',values.MenuSelectPrinter_str);
    end
    % Save as excel
    invoke(Workbook,'SaveAs',outputfile)
    % Close Excel and clean up
    invoke(Excel,'Quit');
    delete(Excel);
    clear Excel;
end

end