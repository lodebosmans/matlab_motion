function createshipmentoverview(condition,shipyear,shipmonth,shipday)

load Temp\values.mat values
if strcmp(condition,'RemoteShipments') == 1 || strcmp(condition,'RemoteShipments_FB') == 1
    readUPSfile(2); % Shipments
end
load Temp\UPSfile_shipment.mat UPSfile_shipment

% Backup the UPS Excel juuuuuuuust in case anyone screws up
%copyfile([values.upsfilepath values.upsfilename],['O:\Administration\BackupFiles\UpsExcel\' values.y values.mo values.d '_' values.h values.mi values.s '_Backup_' values.upsfilename '.xlsx']);

if strcmp(condition,'RegularShipments') == 1
    load Temp\db_production.mat db_production
    load Temp\UPSfile_customer.mat UPSfile_customer %#ok<NASGU>
    load Temp\customers.mat customers
    load Temp\rscases.mat rscases
    
    % Add the ShipTo data
    db_production.sorted  = db_production.overview;
    [nrofpairs,tempcol] = size(db_production.sorted);
    % Get the position of the customer number
    col_cn = catchcolumnindex({'customernumber'},db_production.headers,1);
    col_cn = cell2mat(col_cn(2,1));
    db_production.headers(1,end+1) = {'shipto'};
    db_production.headers(1,end+1) = {'shiptoname'};
    
    col_cstn = catchcolumnindex({'shipto'},db_production.headers,1);
    col_cstn = cell2mat(col_cstn(2,1));
    col_subdelivery = catchcolumnindex({'subdelivery'},db_production.headers,1);
    col_subdelivery = cell2mat(col_subdelivery(2,1));
    col_company = catchcolumnindex({'company'},db_production.headers,1);
    col_company = cell2mat(col_company(2,1));
    col_cstname = catchcolumnindex({'shiptoname'},db_production.headers,1);
    col_cstname = cell2mat(col_cstname(2,1));
    for currentcase = 1:nrofpairs
        % Get the customer number
        cn = char(db_production.sorted(currentcase,col_cn));
        cn = strrep(cn,'-','_');
        %currentcustomer2 = ['c' cn];
        % Get the ShipTo number + name
        db_production.sorted(currentcase,col_cstn) = cellstr(customers.(cn).ShipTo);
        temp_stnr = strrep(char(db_production.sorted(currentcase,col_cstn)),'D','C'); % Convert to 'C' as this is easier for the production operators
        temp_stnr = strrep(temp_stnr,'-','_');
        temp_stname = customers.(temp_stnr).Company;
        db_production.sorted(currentcase,col_cstname) = {temp_stname};
    end
    % Get the column number of the current ship to number (cstn)

    
    % Sort and filter everything
    db_production.sorted = sortrows(db_production.sorted,[col_cstn col_cn]);
    
    %nrofshipments = 0;
    listsendto = {'Empty'};
    liststart = 0;
    % Get the position of the case id
    col_caseid = catchcolumnindex({'caseid.full'},db_production.headers,1);
    col_caseid = cell2mat(col_caseid(2,1));
    
    % Prepare a list of customers for manual selection
    nrofcustomers = size(customers.overview,1);
    list = {};
    list(1) = {'Select an option..'};
    for i = 1:nrofcustomers
        if strcmp(char(customers.overview(i,4)),'Active') == 1
            if length(char(customers.overview(i,1))) == 9
                list(i+1,1) = {[char(customers.overview(i,1)) ' - '  char(customers.overview(i,2)) ' - '  char(customers.overview(i,3))]};
            else
                list(i+1,1) = {[char(customers.overview(i,1)) '         - '  char(customers.overview(i,2)) ' - '  char(customers.overview(i,3))]};
            end
        else
            list(i+1,1) = {'                    - Obsolete'};
        end
    end
    
    for currentcase = 1:nrofpairs
        % Get the (real) customer number
        cstn_c = char(db_production.sorted(currentcase,col_cn));
        cstn_c_ = strrep(cstn_c,'-','_');
        cnshort = cstn_c_(2:5);
        % Get the case id
        ccid = char(db_production.sorted(currentcase,col_caseid)); % ccid = currentcaseid
        ccom = char(db_production.sorted(currentcase,col_company));
        % Check if you need to assign the delivery address manually
        % Only for orders submitted with internal RS Print accounts
        % Exclude the RS Scan account, it will be handled as regular shipments
        if (strcmp(cnshort,'0211') == 1 || strcmp(cnshort,'0210') == 1) && strcmp(cstn_c_,'C0211_006') == 0
            indx = 1;
            % Show error when the default option is still chosen.
            while indx == 1
                promptstring = ['Please find a nice foster home for ' ccid ' ordered by ' ccom];
                [indx,tf] = listdlg('ListString',list,'PromptString',promptstring,'SelectionMode','single','ListSize',[1000,400],'Name','Case ID orphanage','OKString','Assign','CancelString','Do not ship, the destination is not in the list.');
                if indx == 1
                    choice = questdlg('Select something else than the default option.','Caution','Ok','Also ok','Ok');
                    switch choice
                        case 'Ok'
                            donothing = 1;
                        case 'Also ok'
                            donothing = 1;
                    end
                end
                if indx ~= 1
                    indx2 = indx - 1;
                    shiptochoice = [char(customers.overview(indx2,2)) ' - '  char(customers.overview(indx2,3))];
                    message = ['Are you sure you want to ship ' ccid ' to ' shiptochoice '?' ];
                    choice = questdlg(message,'Caution','Ok','No','Ok');
                    switch choice
                        case 'Ok'
                            % Overwrite the customer information to the newly selected information
                            cstn_c = char(customers.overview(indx2,1));
                            db_production.sorted(currentcase,col_cn) = cellstr(cstn_c);
                            % Overwrite the cstn information to the newly selected information
                            cstn_c_ = strrep(cstn_c,'-','_');
                            subdel = eval(['customers.' cstn_c_ '.Subdelivery;']);
                            db_production.sorted(currentcase,col_subdelivery) = cellstr(subdel);
                            cstn_d = eval(['customers.' cstn_c_ '.ShipTo;']);
                            db_production.sorted(currentcase,col_cstn) = cellstr(cstn_d);
                            cstn_d_ = strrep(cstn_d,'-','_');
                            db_production.sorted(currentcase,col_cstname) = {char(customers.overview(indx2,3))};
                            cstname = char(db_production.sorted(currentcase,col_cstname));
                        case 'No'
                            indx = 1;
                    end
                end
                % Remove the case if there is no destination present yet
                if tf == 0
                    indx = 0;
                    cstn_d = 'D0000';
                    cstn_d_ = strrep(cstn_d,'-','_');
                    % remove from list
                    YYYYYYYYYYY = 1;
                    % add to list of cases which are removed
                    YYYYYYYYYYYYYYyy = 1;
                end
            end
            
            
            % if not, assign automatically
        else
            % Get the shipto number
            cstn_d = char(db_production.sorted(currentcase,col_cstn)); % cstn = currentshiptonumber
            cstn_d_ = strrep(cstn_d,'-','_');
            cstname = char(db_production.sorted(currentcase,col_cstname));
        end
        % Do the rest of the general tasks, which are similar.
        
        % Get the (possibly overwritten) customer number
        cstn_c = char(db_production.sorted(currentcase,col_cn));
        cstn_c_ = strrep(cstn_c,'-','_');
        
        % Check if it is a subdelivery or not
        if strcmp(char(db_production.sorted(currentcase,col_subdelivery)),'Yes') == 1
            subdelivery = 1;
        else
            subdelivery = 0;
        end
        
        % Check if the sendto number is already known
        if ismember(cstn_d_,listsendto) == 0
            % Create new shipment
            %nrofshipments = nrofshipments + 1;
            %sizecol_lst = size(listsendto,2);
            if strcmp(char(listsendto(1,1)),'Empty') == 1
                listsendto(numel(listsendto)) = {cstn_d_};
            else
                listsendto(numel(listsendto)+1) = {cstn_d_};
            end
            % Create some variables to set everything up
            eval(['shipments.shipment_' cstn_d_ '.deliverynumber = cstn_d;']);
            eval(['shipments.shipment_' cstn_d_ '.deliveryname = cstname;']);
            eval(['shipments.shipment_' cstn_d_ '.printed = 0;']);
            eval(['shipments.shipment_' cstn_d_ '.caseids = {};']);
            % Add the case ID to the shipment
            eval(['shipments.shipment_' cstn_d_ '.caseids(1,1) = {ccid};']);
            % Add the case ID to subdelivery if necessary
            if subdelivery == 1
                eval(['shipments.shipment_' cstn_d_ '.subdeliveries = struct;']);
                if isfield(eval(['shipments.shipment_' cstn_d_ '.subdeliveries']),cstn_c_) == 0 % Is actually not necessary to add this filter.
                    % Create new subdelivery based on customernumber
                    eval(['shipments.shipment_' cstn_d_ '.subdeliveries.' cstn_c_ ' = {ccid};']);
                end
            end
        else
            % Add to existing shipment
            eval(['shipments.shipment_' cstn_d_ '.caseids(end+1,1) = {ccid};']);
            if subdelivery == 1
                % Check if the subdelivery is already known
                if isfield(eval(['shipments.shipment_' cstn_d_ '.subdeliveries']),cstn_c_) == 0
                    % Create new subdelivery based on customernumber
                    eval(['shipments.shipment_' cstn_d_ '.subdeliveries.' cstn_c_ ' = {ccid};']);
                else
                    % Add to existing subdelivery
                    eval(['shipments.shipment_' cstn_d_ '.subdeliveries.' cstn_c_ '(end+1,1) = {ccid};']);
                end
            end
        end
    end
    save Temp\shipments.mat shipments
elseif strcmp(condition,'RemoteShipments') == 1 || strcmp(condition,'RemoteShipments_FB') == 1
    load Temp\shipments.mat shipments
else
    error('Shipment condition is not known!');
end


if values.NewCaseIDsPresent == 1 || strcmp(condition,'RemoteShipments') == 1 || strcmp(condition,'RemoteShipments_FB') == 1
    % Analyze the amount of shipments
    fldnms = fieldnames(shipments);
    nrofshipments = numel(fldnms);
    
    % Find the row where to write the first shipment (this is not caculated in the Excel file itself, but in the equivalent variable).
    col_year = catchcolumnindex({'Year'},UPSfile_shipment,1);
    col_year = cell2mat(col_year(2,1));
    col_yearxls = char(xlsColNum2Str(col_year));
    col_month = catchcolumnindex({'Month'},UPSfile_shipment,1);
    col_month = cell2mat(col_month(2,1));
    col_monthxls = char(xlsColNum2Str(col_month));
    col_day = catchcolumnindex({'Day'},UPSfile_shipment,1);
    col_day = cell2mat(col_day(2,1));
    col_dayxls = char(xlsColNum2Str(col_day));
    col_shipmentnumber = catchcolumnindex({'ShipmentNumber'},UPSfile_shipment,1);
    col_shipmentnumber = cell2mat(col_shipmentnumber(2,1));
    col_shipmentnumberxls = char(xlsColNum2Str(col_shipmentnumber));
    col_location = catchcolumnindex({'ShippedFrom'},UPSfile_shipment,1);
    col_location = cell2mat(col_location(2,1));
    col_locationxls = char(xlsColNum2Str(col_location));
    col_service = catchcolumnindex({'Service'},UPSfile_shipment,1);
    col_service = cell2mat(col_service(2,1));
    col_servicexls = char(xlsColNum2Str(col_service));
    col_customernr = catchcolumnindex({'CustomerNumber'},UPSfile_shipment,1);
    col_customernr = cell2mat(col_customernr(2,1));
    col_customernrxls = char(xlsColNum2Str(col_customernr));
    col_serialnrs = catchcolumnindex({'SerialNumbers'},UPSfile_shipment,1);
    col_serialnrs = cell2mat(col_serialnrs(2,1));
    col_serialnrsxls = char(xlsColNum2Str(col_serialnrs));
    % col_itemnrs = catchcolumnindex({'ItemNumbers'},UPSfile_shipment,1);
    % col_itemnrs = cell2mat(col_itemnrs(2,1));
    % col_itemnrsxls = char(xlsColNum2Str(col_itemnrs));
    col_promotedoms = catchcolumnindex({'PromotedOMS'},UPSfile_shipment,1);
    col_promotedoms = cell2mat(col_promotedoms(2,1));
    col_promotedomsxls = char(xlsColNum2Str(col_promotedoms));
    col_reference = catchcolumnindex({'Reference'},UPSfile_shipment,1);
    col_reference = cell2mat(col_reference(2,1));
    col_referencesxls = char(xlsColNum2Str(col_reference));
    col_trackingnr = catchcolumnindex({'Trackingnumber'},UPSfile_shipment,1);
    col_trackingnr = cell2mat(col_trackingnr(2,1));
    col_trackingnrxls = char(xlsColNum2Str(col_trackingnr));
    
    lastrowdate = 0;
    counter = 0;
    while lastrowdate == 0
        counter = counter + 1;
        if strcmp(UPSfile_shipment(counter,col_year),'-') == 1
            lastrowdate = counter-1;
        end
    end
    % Now go back to find the last shipment number
    col_sn = catchcolumnindex({'ShipmentNumber'},UPSfile_shipment,1);
    col_sn = cell2mat(col_sn(2,1));
    lastrowshipnum = 0;
    while lastrowshipnum == 0
        if strcmp(UPSfile_shipment(counter,col_sn),'-') == 0
            lastrowshipnum = counter;
        end
        counter = counter - 1;
    end
    lastshipmentline = lastrowdate + 1;
    initialshipmentline = lastshipmentline;
    lastshipmentnr = cell2mat(UPSfile_shipment(lastrowshipnum,col_shipmentnumber)) + 1;
    
    % Open the Excel file and add the data
    % ------------------------------------
    disp('Writing shipments to UPS Excel file - please wait');
    % Specify file name
    file = [values.upsfilepath values.upsfilename]; % This must be full path name
    % Open Excel Automation server
    Excel = actxserver('Excel.Application');
    Workbooks = Excel.Workbooks;
    % Make Excel visible
    if strcmp(condition,'RemoteShipments') == 1 || strcmp(condition,'RemoteShipments_FB') == 1
        Excel.Visible = 1;
    else
        Excel.Visible = 1;
    end
    % Open Excel file
    Workbook=Workbooks.Open(file);
    % Specify sheet number, data, and range to write to
    sheetnum=2;
    % data=rand(4);  % use a cell array if you want both numeric and text data
    % range = 'F10:I13';
    % Make the first sheet active
    Sheets = Excel.ActiveWorkBook.Sheets;
    sheet1 = get(Sheets, 'Item', sheetnum);
    invoke(sheet1, 'Activate');
    Activesheet = Excel.Activesheet;
    
    % For every shipment
    for cshipm = 1:nrofshipments
        % Get fieldname
        temp = char(fldnms(cshipm));
        ccus = ['C' temp(11:end)];
        ccus = strrep(ccus,'_','-');
        % There might need to be some exceptions added here..
        
        % Write the customer number of the main shipment
        range = [col_customernrxls num2str(lastshipmentline) ':' col_customernrxls num2str(lastshipmentline)];
        data = ccus;
        activexwritetoxls(range,data,Activesheet);
        
        % Write the shipmentnumber if applicable
        range = [col_shipmentnumberxls num2str(lastshipmentline) ':' col_shipmentnumberxls num2str(lastshipmentline)];
        data = lastshipmentnr;
        activexwritetoxls(range,data,Activesheet);
        
        % Write the case IDs from the main shipment
        range = [col_serialnrsxls num2str(lastshipmentline) ':' col_serialnrsxls num2str(lastshipmentline)];
        [data] = concatcaseids(shipments.(temp).caseids);
        activexwritetoxls(range,data,Activesheet);
        
        if strcmp(condition,'RemoteShipments_FB') == 1
            range = [col_referencesxls num2str(lastshipmentline) ':' col_referencesxls num2str(lastshipmentline)];
            [data] = shipments.(temp).shipmentnumber;
            activexwritetoxls(range,data,Activesheet);
            
            range = [col_trackingnrxls num2str(lastshipmentline) ':' col_trackingnrxls num2str(lastshipmentline)];
            [data] = shipments.(temp).trackingnumber;
            activexwritetoxls(range,data,Activesheet);
            
        end
        
        lastshipmentline = lastshipmentline + 1;
        
        % Check if there are subdeliveries and add the data
        if isfield(eval(['shipments.' temp ]),'subdeliveries') == 1
            fldnmssubdel = fieldnames(eval(['shipments.' temp '.subdeliveries']));
            nrofsubshipments = numel(fldnmssubdel);
            for csubdel = 1:nrofsubshipments
                % Get fieldname
                ccus = char(fldnmssubdel(csubdel));
                % Write the case IDs from the sub shipment
                range = [col_serialnrsxls num2str(lastshipmentline) ':' col_serialnrsxls num2str(lastshipmentline)];
                [data] = concatcaseids(shipments.(temp).subdeliveries.(ccus));
                activexwritetoxls(range,data,Activesheet);
                % Convert the special sign
                ccus = strrep(ccus,'_','-');
                % Write the customer number of the sub shipment
                range = [col_customernrxls num2str(lastshipmentline) ':' col_customernrxls num2str(lastshipmentline)];
                data = ccus;
                activexwritetoxls(range,data,Activesheet);
                
                lastshipmentline = lastshipmentline + 1;
            end
        end
        lastshipmentnr = lastshipmentnr + 1;
    end
    
    % Write the date
    if strcmp(condition,'RegularShipments') == 1
        % Year
        range = [col_yearxls num2str(initialshipmentline) ':' col_yearxls num2str(lastshipmentline-1)];
        data = ['20' values.y];
        activexwritetoxls(range,data,Activesheet);
        % Month
        range = [col_monthxls num2str(initialshipmentline) ':' col_monthxls num2str(lastshipmentline-1)];
        data = values.mo;
        activexwritetoxls(range,data,Activesheet);
        % Day
        range = [col_dayxls num2str(initialshipmentline) ':' col_dayxls num2str(lastshipmentline-1)];
        data = values.d;
        activexwritetoxls(range,data,Activesheet);
    elseif strcmp(condition,'RemoteShipments') == 1 || strcmp(condition,'RemoteShipments_FB') == 1
        % Year
        range = [col_yearxls num2str(initialshipmentline) ':' col_yearxls num2str(lastshipmentline-1)];
        data = shipyear;
        activexwritetoxls(range,data,Activesheet);
        % Month
        range = [col_monthxls num2str(initialshipmentline) ':' col_monthxls num2str(lastshipmentline-1)];
        data = shipmonth;
        activexwritetoxls(range,data,Activesheet);
        % Day
        range = [col_dayxls num2str(initialshipmentline) ':' col_dayxls num2str(lastshipmentline-1)];
        data = shipday;
        activexwritetoxls(range,data,Activesheet);
        % Promoted in OMS
        range = [col_promotedomsxls num2str(initialshipmentline) ':' col_promotedomsxls num2str(lastshipmentline-1)];
        data = '1';
        activexwritetoxls(range,data,Activesheet);
    end
    % Location
    range = [col_locationxls num2str(initialshipmentline) ':' col_locationxls num2str(lastshipmentline-1)];
    if strcmp(condition,'RemoteShipments_FB') == 1
        data = 'FB';
    else
        data = values.CountryOfShipment;
    end
    activexwritetoxls(range,data,Activesheet);
    % service
    range = [col_servicexls num2str(initialshipmentline) ':' col_servicexls num2str(lastshipmentline-1)];
    if strcmp(condition,'RemoteShipments_FB') == 1
        data = 'FedEx';
    else
        data = 'UPS';
    end
    activexwritetoxls(range,data,Activesheet);
    
    % Overrule UPS shipment to Runners' Lab. Go over all the inserted lines and overwrite if necessary.
    for cline = initialshipmentline:lastshipmentline-1
        % Get content of the cell
        Range = get(Activesheet,'range',[col_customernrxls num2str(cline)]);
        ccus = Range.value;
        if strcmp(ccus,'C1000-002') == 1 || values.ManDelNote == 1
            activexwritetoxls([col_servicexls num2str(cline) ':' col_servicexls num2str(cline)],values.Username_str,Activesheet);
        end
    end
    
    if strcmp(condition,'RegularShipments') == 1
        % Double check is the data is correct.
        message = 'Please check if all inserted data in the UPS Excel is correct. If necessary, add manual items. If everything is correct, press ''Ok'' to continue. Press ''Cancel'' to abort and close the Excel without saving.';
        choice = questdlg(message,'Caution','Ok','Cancel','Ok');
        switch choice
            case 'Ok'
                donothing = 1;
                % Save file
                invoke(Workbook,'Save')
                % Close Excel and clean up
                invoke(Excel,'Quit');
                delete(Excel);
                clear Excel;
            case 'Cancel'
                % Close Excel and clean up
                Excel.DisplayAlerts = 0;
                invoke(Excel,'Quit');
                delete(Excel);
                clear Excel;
                clear all
        end
    elseif strcmp(condition,'RemoteShipments') == 1 || strcmp(condition,'RemoteShipments_FB') == 1
        invoke(Workbook,'Save')
        % Close Excel and clean up
        invoke(Excel,'Quit');
        delete(Excel);
        clear Excel;
    else
        error('Not good..')
    end
end

end
