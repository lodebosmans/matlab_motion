function createnavisionshipmentinput()

load Temp\values.mat values
load Temp\UPSfile_shipment.mat UPSfile_shipment
%load Temp\UPSfileUS.mat UPSfileUS
load Temp\customers customers  %#ok<NASGU>
copyfile([values.salesordercheck_folder 'salesordercheck_v3.mat'],'Temp\salesordercheck_v3.mat'); % always keep full path
load Temp\salesordercheck_v3.mat salesordercheck_v3

disp('Creating Navision shipment input - please wait')

% Open file to write all sales orders that are processed in this run
fid = fopen(['Output\Navision\' values.y values.mo values.d '_' values.h values.mi values.s '_SS_XMLsPrepared.txt'],'wt');

% Determine the end date. If end date empty, equal to start. If not, take end date.
possibledays = 1:31;
% Write a message with the range of the dates.
startdate = ['01/' values.RequestedMonth '/' values.RequestedYear];
enddate = ['31/' values.RequestedMonth '/' values.RequestedYear];
message = ['Sales shipments from ' values.RequestedMonth '/' values.RequestedYear ' are being generated.' ];
disp(message);
disp(' ');
fprintf(fid,message);
fprintf(fid, '\n');

% Take backup of salesordercheck variable.
copyfile([values.salesordercheck_folder 'salesordercheck_v3.mat'],['Output\Navision\' values.y values.mo values.d '_' values.h values.mi values.s '_pre_salesordercheck_v3.mat']);

shipmentcounter = 0;
caseidcounter = 0;

% For the two different files (BE/US)
%labelsBE = {'ShipmentNumber','Trackingnumber','Year','Month','Day','Service','SerialNumbers'};
%labelsUS = {'ShipmentNumber','Trackingnumber','Year','Month','Day','Service','Insoles'};
for x = 1:1
    % 1 = BE
    % 2 = US - FB
    % Get the correct label set
    if x == 1
        %labels = labelsBE;
        upsfile = UPSfile_shipment;
        col_SerialNumbers = catchcolumnindex({'SerialNumbers'},upsfile,1);
        col_SerialNumbers = cell2mat(col_SerialNumbers(2,1));
        col_Service = catchcolumnindex({'Service'},upsfile,1);
        col_Service = cell2mat(col_Service(2,1));
%         location = 'PRODBE';
    elseif x == 2
        %labels = labelsUS;
        upsfile = UPSfileUS;
        col_SerialNumbers = catchcolumnindex({'Insoles'},upsfile,1);
        col_SerialNumbers = cell2mat(col_SerialNumbers(2,1));
        service = 'UPS';
%         location = 'PRODUS';
    end
    % Get the columnnumbers from the UPS file
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
    col_ShippedFrom = catchcolumnindex({'ShippedFrom'},upsfile,1);
    col_ShippedFrom = cell2mat(col_ShippedFrom(2,1));
    
    % Create the overview for Flowbuilt/Superfeet and GO4D
    distributors = {'go4d','superfeet'};
    distributors_code = {'C0022','C0032'};
    invoice_overview_columns = {'CustomerNumber','CompanyName','OfficeName','CaseID','Price','ItemNumber','Description','Reference','TrackingNumber','ShipmentDate','Agent'};
    for cpos = 1:size(distributors,2)
        cdistri_name = char(distributors(1,cpos));
        
        eval(['invoice_overview_' cdistri_name ' = invoice_overview_columns;']);
        
        if cpos == 1
            col_invoiceoverview_CustomerNumber = catchcolumnindex({'CustomerNumber'},invoice_overview_columns,1);
            col_invoiceoverview_CustomerNumber = cell2mat(col_invoiceoverview_CustomerNumber(2,1));
            col_invoiceoverview_CompanyName = catchcolumnindex({'CompanyName'},invoice_overview_columns,1);
            col_invoiceoverview_CompanyName = cell2mat(col_invoiceoverview_CompanyName(2,1));
            col_invoiceoverview_OfficeName = catchcolumnindex({'OfficeName'},invoice_overview_columns,1);
            col_invoiceoverview_OfficeName = cell2mat(col_invoiceoverview_OfficeName(2,1));
            col_invoiceoverview_CaseID = catchcolumnindex({'CaseID'},invoice_overview_columns,1);
            col_invoiceoverview_CaseID = cell2mat(col_invoiceoverview_CaseID(2,1));
            col_invoiceoverview_Price = catchcolumnindex({'Price'},invoice_overview_columns,1);
            col_invoiceoverview_Price = cell2mat(col_invoiceoverview_Price(2,1));
            col_invoiceoverview_ItemNumber = catchcolumnindex({'ItemNumber'},invoice_overview_columns,1);
            col_invoiceoverview_ItemNumber = cell2mat(col_invoiceoverview_ItemNumber(2,1));
            col_invoiceoverview_Description = catchcolumnindex({'Description'},invoice_overview_columns,1);
            col_invoiceoverview_Description = cell2mat(col_invoiceoverview_Description(2,1));
            col_invoiceoverview_Reference = catchcolumnindex({'Reference'},invoice_overview_columns,1);
            col_invoiceoverview_Reference = cell2mat(col_invoiceoverview_Reference(2,1));
            col_invoiceoverview_Trackingnumber = catchcolumnindex({'TrackingNumber'},invoice_overview_columns,1);
            col_invoiceoverview_Trackingnumber = cell2mat(col_invoiceoverview_Trackingnumber(2,1));
            col_invoiceoverview_ShipmentDate = catchcolumnindex({'ShipmentDate'},invoice_overview_columns,1);
            col_invoiceoverview_ShipmentDate = cell2mat(col_invoiceoverview_ShipmentDate(2,1));
            col_invoiceoverview_Agent = catchcolumnindex({'Agent'},invoice_overview_columns,1);
            col_invoiceoverview_Agent = cell2mat(col_invoiceoverview_Agent(2,1));
        end
    end
    
    % SalesOrderCheck header columns
    col_SOC_CaseID = catchcolumnindex({'CaseID'},salesordercheck_v3.headers,1);
    col_SOC_CaseID = cell2mat(col_SOC_CaseID(2,1));
    col_SOC_Price = catchcolumnindex({'Price'},salesordercheck_v3.headers,1);
    col_SOC_Price = cell2mat(col_SOC_Price(2,1));
    col_SOC_CustomerNumber = catchcolumnindex({'CustomerNumber'},salesordercheck_v3.headers,1);
    col_SOC_CustomerNumber = cell2mat(col_SOC_CustomerNumber(2,1));
    col_SOC_CompanyName = catchcolumnindex({'CompanyName'},salesordercheck_v3.headers,1);
    col_SOC_CompanyName = cell2mat(col_SOC_CompanyName(2,1));
    col_SOC_DeliveryFacility = catchcolumnindex({'DeliveryFacility'},salesordercheck_v3.headers,1);
    col_SOC_DeliveryFacility = cell2mat(col_SOC_DeliveryFacility(2,1));
    col_SOC_ItemNumber = catchcolumnindex({'ItemNumber'},salesordercheck_v3.headers,1);
    col_SOC_ItemNumber = cell2mat(col_SOC_ItemNumber(2,1));
    
    % SalesOrderCheck DataFeed columns
    col_DF_CaseID = catchcolumnindex({'CaseCode'},salesordercheck_v3.datafeed,1);
    col_DF_CaseID = cell2mat(col_DF_CaseID(2,1));
    col_DF_ServiceType = catchcolumnindex({'ServiceType'},salesordercheck_v3.datafeed,1);
    col_DF_ServiceType = cell2mat(col_DF_ServiceType(2,1));
    col_DF_InsoleType = catchcolumnindex({'InsoleType'},salesordercheck_v3.datafeed,1);
    col_DF_InsoleType = cell2mat(col_DF_InsoleType(2,1));
    col_DF_Top_thickness = catchcolumnindex({'Top_thickness'},salesordercheck_v3.datafeed,1);
    col_DF_Top_thickness = cell2mat(col_DF_Top_thickness(2,1));
    col_DF_Top_size = catchcolumnindex({'Top_size'},salesordercheck_v3.datafeed,1);
    col_DF_Top_size = cell2mat(col_DF_Top_size(2,1));
    col_DF_Top_hardness = catchcolumnindex({'Top_hardness'},salesordercheck_v3.datafeed,1);
    col_DF_Top_hardness = cell2mat(col_DF_Top_hardness(2,1));
    col_DF_Top_material = catchcolumnindex({'Top_material'},salesordercheck_v3.datafeed,1);
    col_DF_Top_material = cell2mat(col_DF_Top_material(2,1));
    col_DF_Base_type = catchcolumnindex({'Base_type'},salesordercheck_v3.datafeed,1);
    col_DF_Base_type = cell2mat(col_DF_Base_type(2,1));
    col_DF_MeasurementUnit = catchcolumnindex({'MeasurementUnit'},salesordercheck_v3.datafeed,1);
    col_DF_MeasurementUnit = cell2mat(col_DF_MeasurementUnit(2,1));
    
    % Compensate for the bad coding that I put a cell in a cell, which is
    % not readable by the catchrowindex code..
    caseids_DF = cell(0);
    for cdf = 1 : size(salesordercheck_v3.datafeed,1)
        temp = salesordercheck_v3.datafeed(cdf,1);
        caseids_DF(end+1,1) = cellstr(char(temp{1}));
    end
    
    invoice_overview = cell(0);
    
    
    % Go over all the rows in the shipment excel
    % If dates matches, continue
    nrofrows = size(upsfile,1);
    % Skip the first row with the headers
    for cr = 2:nrofrows
        % If the date of shipment matches the selected date's and the shipmentnumber is not empty, continue.
        if cell2mat(upsfile(cr,col_Year)) == str2double(values.RequestedYear) ...
                && cell2mat(upsfile(cr,col_Month)) == str2double(values.RequestedMonth) ...
                && ismember(cell2mat(upsfile(cr,col_Day)),possibledays) ...
                && strcmp(cell2mat(upsfile(cr,col_ShipmentNumber)),'-') == 0
            
            shipmentcounter = shipmentcounter + 1;
            % Define the shipment date
            shipyear = num2str(cell2mat(upsfile(cr,col_Year)));
            if cell2mat(upsfile(cr,col_Month)) < 10
                shipmonth = ['0' num2str(cell2mat(upsfile(cr,col_Month)))];
            else
                shipmonth = num2str(cell2mat(upsfile(cr,col_Month)));
            end
            if cell2mat(upsfile(cr,col_Day)) < 10
                shipday = ['0' num2str(cell2mat(upsfile(cr,col_Day)))];
            else
                shipday = num2str(cell2mat(upsfile(cr,col_Day)));
            end
            shipmentdate = [shipyear '/' shipmonth '/' shipday];
            % Get the case IDs from that shipment
            currentshipment = char(upsfile(cr,col_SerialNumbers));
            %[caseids] = readmanualcaseids2(values,currentshipment);
            [caseids] = readmanualcaseids({},'',0,currentshipment,0);
            
            tracknrtemp = upsfile(cr,col_Trackingnumber);
            if iscellstr(tracknrtemp) == 1
                tracknrtemp = char(tracknrtemp);
            else
                tracknrtemp = num2str(cell2mat(tracknrtemp));
            end
            if isnumeric(tracknrtemp) == 1
                tracknrtemp = num2str(tracknrtemp);
            end
            
            if contains(tracknrtemp,'/') == 1
                positionslash = find(tracknrtemp=='/');
                tracknrtemp = tracknrtemp(1:positionslash(1,1)-1); % (1,1) in case there is more than one slash present
            end
            
            location = char(upsfile(cr,col_ShippedFrom));
            reference = char(upsfile(cr,col_Reference));
            
            nrofcasesinshipment = size(caseids,1);
            for cc = 1:nrofcasesinshipment
                caseidcounter = caseidcounter + 1;
                % Get the case id
                currentcaseid = char(caseids(cc,1));
                if strcmp(currentcaseid,'RS20-ILU-PAS') == 1
                    test = 1;
                end
                if strcmp(currentcaseid,'RS18-JAZ-NIH') == 0 % was cancelled, but also shipped..
                    % Display a message
                    message = ['Processing case ' currentcaseid ' on row ' num2str(cr) ' of ' location ' UPS file (shipment ' num2str(cell2mat(upsfile(cr,col_ShipmentNumber))) '), '];
                    disp(message);
                    fprintf(fid,message);
                    if length(currentcaseid) == 12
                        
                        % Find the corresponding row in the salesordercheck variable
                        [temp] = catchrowindex({currentcaseid},salesordercheck_v3.data,1);
                        rownr = cell2mat(temp(2,1)); % If something goes wrong here, the sales order linked to this case ID is not yet uploaded in Navision.
                        % Check if the case ID is already shipped, if so, add additional date in last column. Do not overwrite previous data
                        if strcmp(char(salesordercheck_v3.data(rownr,10)),'-') == 1
                            % It is the first shipment of this case
                            % Insert (1) the production site, (2) the shipment date and (3) the tracking number.
                            salesordercheck_v3.data(rownr,10) = {location};
                            salesordercheck_v3.data(rownr,11) = {shipmentdate};
                            if x == 1
                                salesordercheck_v3.data(rownr,12) = upsfile(cr,col_Service);
                            elseif x == 2
                                salesordercheck_v3.data(rownr,12) = {service};
                            end
                            salesordercheck_v3.data(rownr,13) = {tracknrtemp};
                            salesordercheck_v3.data(rownr,14) = upsfile(cr,col_Reference);
                            message = ['which is linked to row ' num2str(rownr) ' in the onlinereport/salesordercheck_v3.'];
                            disp(message);
                            disp(' ');
                            fprintf(fid,message);
                            fprintf(fid, '\n');
                        else
                            % The case has been shipped before
                            reshipment = ['(' location '/' shipmentdate '/' tracknrtemp '/' reference ')///'];
                            salesordercheck_v3.data(rownr,15) = {[ char(salesordercheck_v3.data(rownr,15)) '///' reshipment]};
                        end
                         
                    elseif length(currentcaseid) == 13
                        rownr = 'NoNumber';
                    end
                    
                    if strcmp(location,'BE') == 1
                        % Write the god damned sales shipment xml
                        % ---------------------------------------
                        
                        docNode = com.mathworks.xml.XMLUtils.createDocument('SalesOrdersToShip');
                        docRootNode = docNode.getDocumentElement;
                        
                        SalesOrder_tag = docNode.createElement('SalesOrder');
                        docRootNode.appendChild(SalesOrder_tag);
                        
                        CaseID_tag = docNode.createElement('CaseID');
                        CaseID_text = docNode.createTextNode(currentcaseid);
                        CaseID_tag.appendChild(CaseID_text);
                        SalesOrder_tag.appendChild(CaseID_tag);
                        
                        PackageTrackingCode_tag = docNode.createElement('PackageTrackingCode');
                        PackageTrackingCode_text = docNode.createTextNode(tracknrtemp);
                        PackageTrackingCode_tag.appendChild(PackageTrackingCode_text);
                        SalesOrder_tag.appendChild(PackageTrackingCode_tag);
                        
                        PostingDate_tag = docNode.createElement('PostingDate');
                        PostingDate_text = docNode.createTextNode(shipmentdate);
                        PostingDate_tag.appendChild(PostingDate_text);
                        SalesOrder_tag.appendChild(PostingDate_tag);
                        
                        % Write to file
                        xmlFileName = ['Output\Navision\SalesShipment' num2str(rownr) '_' currentcaseid '.xml'];
                        xmlwrite(xmlFileName,docNode);
                        xmlFileName = [values.navfolder  'SalesShipments\SalesShipment' num2str(rownr) '_' currentcaseid '.xml'];
                        xmlwrite(xmlFileName,docNode);
                        %type(xmlFileName);
                        
                    elseif strcmp(location,'FB') == 1
                        
                        % Search the case in the SOC
                        row_SOC_CaseID = catchrowindex({currentcaseid},salesordercheck_v3.data,col_SOC_CaseID);
                        row_SOC_CaseID = cell2mat(row_SOC_CaseID(2,1));
                        
                        % Find the corresponding row in the datafeed variable
                        row_DF_CaseID = catchrowindex({currentcaseid},caseids_DF,col_DF_CaseID);
                        row_DF_CaseID = cell2mat(row_DF_CaseID(2,1));
                        
                        % Get the data for the nice description
                        label = ["ServiceType","InsoleType","Top_thickness","Top_size","Top_hardness","Top_material","Base_type","MeasurementUnit"];
                        for ctext = 1:size(label,2)
                            clabel = char(label(1,ctext));
                            eval(['temp = salesordercheck_v3.datafeed(row_DF_CaseID,col_DF_' clabel ');']);
                            eval(['label_' clabel ' = char(temp{1});']);
                        end
                        
                        description = [label_InsoleType ' ' label_Base_type ' ' label_Top_size ' ' label_MeasurementUnit ' ' label_Top_material ' ' label_Top_thickness ' ' label_Top_hardness ' ' label_ServiceType ];
                        
                        ccnr = char(salesordercheck_v3.data(row_SOC_CaseID,col_SOC_CustomerNumber));
                        ccnr_ = strrep(ccnr,'-','_');
                        
%                         customers.(ccnr_).Agent
                        
                        nextrow = size(invoice_overview,1) + 1;
                        
                        invoice_overview(nextrow,col_invoiceoverview_CaseID) = cellstr(currentcaseid);
                        invoice_overview(nextrow,col_invoiceoverview_CompanyName) = salesordercheck_v3.data(row_SOC_CaseID,col_SOC_CompanyName);
                        invoice_overview(nextrow,col_invoiceoverview_CustomerNumber) = cellstr(ccnr);
                        invoice_overview(nextrow,col_invoiceoverview_OfficeName) = salesordercheck_v3.data(row_SOC_CaseID,col_SOC_DeliveryFacility);
                        invoice_overview(nextrow,col_invoiceoverview_Price) = salesordercheck_v3.data(row_SOC_CaseID,col_SOC_Price);
                        invoice_overview(nextrow,col_invoiceoverview_ItemNumber) = salesordercheck_v3.data(row_SOC_CaseID,col_SOC_ItemNumber);
                        invoice_overview(nextrow,col_invoiceoverview_Reference) = cellstr(reference);
                        invoice_overview(nextrow,col_invoiceoverview_Description) = cellstr(description);
                        invoice_overview(nextrow,col_invoiceoverview_ShipmentDate) = cellstr(shipmentdate);
                        invoice_overview(nextrow,col_invoiceoverview_Trackingnumber) = cellstr(tracknrtemp);
                        invoice_overview(nextrow,col_invoiceoverview_Agent) = cellstr(customers.(ccnr_).Agent);
                        
                    end
                end
            end
        end
    end
end

test = 1;
% Sort the rows based on the agent number
invoice_overview = sortrows(invoice_overview,col_invoiceoverview_Agent);
% temp(1,1:size(invoice_overview_columns,2)) = invoice_overview_columns;
% temp(2:size(invoice_overview,1)+1,:) = invoice_overview;
% invoice_overview = temp;
% clear temp;

for d = 1:size(distributors,2)
    cdis_nr = char(distributors_code(d));
    cdis_name = char(distributors(d));
    temp = {};
    temp(1,1:size(invoice_overview_columns,2)) = invoice_overview_columns;
    for r = 1:size(invoice_overview,1)
         if contains(invoice_overview(r,col_invoiceoverview_Agent),cdis_nr) == 1
            temp(size(temp,1)+1,:) = invoice_overview(r,:);
         end
    end
    filename = ['invoice_overview_' cdis_name];
    eval([filename ' = temp;']);
    % Write to file
    xlswrite(['Output\Navision\' values.y values.mo values.d '_' values.h values.mi values.s '_' cdis_name '_InvoiceOverview.xlsx'],eval(filename));
end

if shipmentcounter == 0 || caseidcounter == 0
    if shipmentcounter == 0
        message = 'No shipments were found on the selected days.';
    elseif caseidcounter == 0
        message = 'Shipments were found on the selected days, however containing no case IDs.';
    end
    disp(message);
    fprintf(fid,message);
    fprintf(fid, '\n');
end
% save the file (in regular interface folder - not in test folder)
save([values.salesordercheck_folder '\salesordercheck_v3.mat'],'salesordercheck_v3');
save(['Output\Navision\' values.y values.mo values.d '_' values.h values.mi values.s '_values.mat'],'values');
copyfile([values.salesordercheck_folder 'salesordercheck_v3.mat'],['Output\Navision\' values.y values.mo values.d '_' values.h values.mi values.s '_post_salesordercheck_v3.mat']);
% Save as excel as well
xlswrite([values.salesordercheck_folder 'salesordercheck_v3.xlsx'],salesordercheck_v3.data);
% xlswrite(['Output\Navision\' values.y values.mo values.d '_' values.h values.mi values.s  ' GO4D_InvoiceOverview.xlsx'],invoice_overview);

fclose(fid);

% salesordercheck_v3(:,4:9) = {' '};

end
