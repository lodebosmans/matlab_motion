function createdeliverynote_format2019(condition)

load Temp\values.mat values
load Temp\UPSfile_shipment.mat UPSfile_shipment
load Temp\customers.mat customers
copyfile([values.salesordercheck_folder 'salesordercheck_v3.mat'],'Temp\salesordercheck_v3.mat'); % always keep full pathway
load Temp\salesordercheck_v3.mat salesordercheck_v3
load Temp\onlinereport.mat onlinereport

clc
disp('Creating delivery notes - please wait')

info.shipmentcounter = 0;
info.specialsthingstodo = cell(0);
info.specialsamount = size(info.specialsthingstodo,1);

% Get the column numbers.
col_shipped = catchcolumnindex({'PromotedOMS'},UPSfile_shipment,1);
col_shipped = cell2mat(col_shipped(2,1));
col_service = catchcolumnindex({'Service'},UPSfile_shipment,1);
col_service = cell2mat(col_service(2,1));
%     col_shipnr = catchcolumnindex({'ShipmentNumber'},UPSfile_shipment,1);
%     col_shipnr = cell2mat(col_shipnr(2,1));
col_shippedFrom = catchcolumnindex({'ShippedFrom'},UPSfile_shipment,1);
col_shippedFrom = cell2mat(col_shippedFrom(2,1));


[nrofrows,nrofcols] = size(UPSfile_shipment); %#ok<ASGLU>
for cr = 2:nrofrows
    if strcmp(condition,'UpsShipment') == 1
        % Check is the shipment in the row is already shipped. If not, fetch all necessary data
        if strcmp(cell2mat(UPSfile_shipment(cr,col_shipped)),'-') == 1 && strcmp(UPSfile_shipment(cr,col_service),'UPS') == 1 && strcmp(char(UPSfile_shipment(cr,col_shippedFrom)),values.CountryOfShipment) == 1
            info.shipmentcounter = info.shipmentcounter + 1;
            [info] = writedeliverynotetofile(info,condition,customers,onlinereport,salesordercheck_v3,UPSfile_shipment,cr);
        end
    elseif strcmp(condition,'ManualShipment') == 1
        % Check is the shipment in the row is already shipped. If not, fetch all necessary data
        if strcmp(cell2mat(UPSfile_shipment(cr,col_shipped)),'-') == 1 && strcmp(UPSfile_shipment(cr,col_service),'UPS') == 0 && strcmp(char(UPSfile_shipment(cr,col_shippedFrom)),values.CountryOfShipment) == 1
            info.shipmentcounter = info.shipmentcounter + 1;
            [info] = writedeliverynotetofile(info,condition,customers,onlinereport,salesordercheck_v3,UPSfile_shipment,cr);
        end
    end    
end



% Print the special things to do
if info.specialsamount > 0
    file = [values.deliverynotetemplatefolder values.shipmentexceptionsfilename]; % This must be full path name  => only adapt the file in the 'Interface' folder, not in the test environment.
    Excel = actxserver('Excel.Application');
    % Set it to visible
    set(Excel,'Visible',0);
    % Add a Workbook
    Workbooks = Excel.Workbooks;
    Workbook = Workbooks.Open(file);
    %Workbook = invoke(Workbooks, 'Add');
    % Get a handle to Sheets and select Sheet 1
    Sheets = Excel.ActiveWorkBook.Sheets;
    Sheet1 = get(Sheets, 'Item', 1);
    Sheet1.Activate;
    eActivesheetRange = get(Excel.Activesheet,'Range','A1:A1');
    eActivesheetRange.Value = [values.d '/' values.mo '/20' values.y ' - Shipment exceptions'];
    eActivesheetRange = get(Excel.Activesheet,'Range',['B3:D' num2str(info.specialsamount+2)]);
    eActivesheetRange.Value = info.specialsthingstodo;
    
    % Delete the unnecessary rows.
    for y = 1:8
        for x = info.specialsamount+3:47
            Sheet1.Rows.Item(x).Delete;
        end
    end
    
    Sheet1.Columns.Item(2).columnWidth = 10;
    Sheet1.Columns.Item(3).columnWidth = 20;
    Sheet1.Columns.Item(4).columnWidth = 50;
    
    % And set it to text wrap
    eActivesheetRange = get(Excel.Activesheet,'Range','B:D');
    eActivesheetRange.WrapText = true;
    % Align everything vertically upwards
    eActivesheetRange = get(Excel.Activesheet,'Range','B:D');
    eActivesheetRange.VerticalAlignment = 1;
    
    % Create the output filename
    outputfile = [pwd '\Output\DeliveryNotes\' values.y values.mo values.d '_' values.h values.mi values.s '_OverviewSpecialThings'];
    
    % Send to printer
    if values.MenuSelectPrinter > 1
        Excel.ActiveWorkbook.PrintOut(1,1,1,'False',values.MenuSelectPrinter_str);
    end
    % Save as excel
    invoke(Workbook,'SaveAs',outputfile);
    % Close Excel and clean up
    invoke(Excel,'Quit');
    delete(Excel);
    clear Excel;
end

end