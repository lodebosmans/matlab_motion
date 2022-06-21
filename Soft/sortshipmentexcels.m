function sortshipmentexcels()

load Temp\values.mat values

disp('Checking if some shipment excels need to be sorted out - please wait');

% Procedure for xlsx files.
files = dir([values.shipmentexcelfolder 'New\*.xlsx']);
nroffiles = length(files);
if nroffiles > 0
    disp(['Sorting out ' num2str(nroffiles) ' shipment excel(s) - please wait']);
    for k = 1:nroffiles
        shipments = struct;
        info.filename = files(k).name;
        disp(['Sorting out: ' info.filename]);
        % Get the year and month
        info.year = info.filename(1:4);
        info.month = info.filename(5:6);
        info.day = info.filename(7:8);
        % Get the destination
        %if strcmp(info.filename(30:31),'US') == 1
        if contains(info.filename,'_US') == 1
            disp('Easy one... straight copy to the US folder.');
            info.destination = 'US';
            info.cstn = 'ignore';
            [info] = addmoreinfo(info,'sortshipmentexcels');
            sortshipmentexceltofolder(info);
        elseif contains(info.filename,'_Paal') == 1
            disp('Easy one... straight copy to the Paal folder.');
            info.destination = 'Paal';
            info.cstn = 'ignore';
            [info] = addmoreinfo(info,'sortshipmentexcels');
            sortshipmentexceltofolder(info);
        elseif contains(info.filename,'_FLOWBUILT') == 1
            disp('Easy one... straight copy to the Flowbuilt folder.');
            info.destination = 'FLOWBUILT';
            info.cstn = 'ignore';
            [info] = addmoreinfo(info,'sortshipmentexcels');
            sortshipmentexceltofolder(info);
        elseif contains(info.filename,'_LIVIT') == 1
            info.destination = 'LIVIT';
            info.cstn = 'D1071';
            [info] = addmoreinfo(info,'sortshipmentexcels');
            [shipments] = getshippedcaseids(info,shipments);
            save Temp\shipments.mat shipments
            if info.alreadypresent == 0
                createshipmentoverview('RemoteShipments',info.year,info.month,info.day);
            end
            sortshipmentexceltofolder(info);
        elseif contains(info.filename,'_RUNLAB') == 1
            info.destination = 'RUNLAB';
            info.cstn = 'D1000';
            [info] = addmoreinfo(info,'sortshipmentexcels');
            [shipments] = getshippedcaseids(info,shipments);
            save Temp\shipments.mat shipments
            if info.alreadypresent == 0
                createshipmentoverview('RemoteShipments',info.year,info.month,info.day);
            end
            sortshipmentexceltofolder(info);
        elseif contains(info.filename,'_RSSCAN') == 1
            info.destination = 'RSSCAN';
            info.cstn = 'D0001';
            [info] = addmoreinfo(info,'sortshipmentexcels');
            [shipments] = getshippedcaseids(info,shipments);
            save Temp\shipments.mat shipments
            if info.alreadypresent == 0
                createshipmentoverview('RemoteShipments',info.year,info.month,info.day);
            end
            sortshipmentexceltofolder(info);
        else
            % Do nothing with this file
            disp(['A destination could not be defined for ' info.filename]);
        end
        clear info
        clear shipments
    end
end



% Procedure for csv files.
files = dir([values.shipmentexcelfolder 'New\*.csv']);
nroffiles = length(files);
if nroffiles > 0
    disp(['Sorting out ' num2str(nroffiles) ' shipment excel(s) - please wait']);
    for k = 1:nroffiles
        shipments = struct;
        info.filename = files(k).name;
        disp(['Sorting out: ' info.filename]);
        % Get the year and month
        info.year = info.filename(1:4);
        info.month = info.filename(5:6);
        info.day = info.filename(7:8);
        
        % Get the destination
        if contains(info.filename,'FB_DEALER') == 1
            info.destination = 'FB_DEALER';
            [info] = addmoreinfo(info,'sortshipmentexcels_FB');
            
            [brol1,brol2,c]= xlsread(info.newfile);
            sizeshipment = size(c,1);
            if sizeshipment > 1
                for cl = 1:sizeshipment
                    cline = char(c(cl,1));
                    temp = strsplit(cline,',');
                    temp_rows = size(temp,2);
                    FB_data(cl,1:temp_rows) = temp;
                end
                
                col_caseid = catchcolumnindex({'Case Code'},FB_data,1);
                col_caseid = cell2mat(col_caseid(2,1));
                col_customernr = catchcolumnindex({'Account Number'},FB_data,1);
                col_customernr = cell2mat(col_customernr(2,1));
                col_shipdate = catchcolumnindex({'Ship Date'},FB_data,1);
                col_shipdate = cell2mat(col_shipdate(2,1));
                col_shipnr = catchcolumnindex({'Shipment Number'},FB_data,1);
                col_shipnr = cell2mat(col_shipnr(2,1));
                col_trackingnr = catchcolumnindex({'Tracking Number'},FB_data,1);
                col_trackingnr = cell2mat(col_trackingnr(2,1));
                
                % Get the shipment date
                date = char(FB_data(cl,col_shipdate));
                date = strrep(date,'  ',' ');
                temp2 = find(date == ' ');
                info.month = date(1:temp2(1,1)-1);
                switch info.month
                    case 'Jan'
                        info.month = '01';
                    case 'Feb'
                        info.month = '02';
                    case 'Mar'
                        info.month = '03';
                    case 'Apr'
                        info.month = '04';
                    case 'May'
                        info.month = '05';
                    case 'Jun'
                        info.month = '06';
                    case 'Jul'
                        info.month = '07';
                    case 'Aug'
                        info.month = '08';
                    case 'Sep'
                        info.month = '09';
                    case 'Oct'
                        info.month = '10';
                    case 'Nov'
                        info.month = '11';
                    case 'Dec'
                        info.month = '12';
                end
                info.day = date(temp2(1,1)+1:temp2(1,2)-1);
                if length(info.day) == 1
                    info.day = ['0' info.day];
                end
                info.year = date(temp2(1,2)+1:temp2(1,3)-1);
                % Run again to correct data.
                [info] = addmoreinfo(info,'sortshipmentexcels_FB');
                
                % Sort the cases per customer.
                nrofFBshipments = 0;
                for cl = 2:sizeshipment
                    
                    % Get the customer number
                    ccnr = FB_data(cl,col_customernr);
                    info.cstn = strrep(ccnr,'C','D');
                    info.cstn = char(strrep(info.cstn,' ','_'));
                    info.cstn = char(strrep(info.cstn,'-','_'));
                    disp(info.cstn);
                    
                    % Get the case ID
                    ccid = FB_data(cl,col_caseid);
                    cshnr = FB_data(cl,col_shipnr);
                    ctrnr = FB_data(cl,col_trackingnr);
                    
                    if strcmp('RS20-MAD-REK',char(ccid)) == 1
                        test = 1;
                    end
                    
                    % Check if the shiptoaddress already exists
                    if isfield(shipments,['shipment_' info.cstn]) == 0
                        % Create new shipment
                        eval(['shipments.shipment_' info.cstn '.deliverynumber = info.cstn;']);
                        % Add the reference
                        eval(['shipments.shipment_' info.cstn '.shipmentnumber = cshnr;']);
                        % Add the tracking number
                        eval(['shipments.shipment_' info.cstn '.trackingnumber = ctrnr;']);
                        % Add the case IDs to the shipment
                        eval(['shipments.shipment_' info.cstn '.caseids = ccid;']);
                        
                    else
                        % Add to existing shipment
                        eval(['shipments.shipment_' info.cstn '.caseids(end+1,1) = ccid;']);
                    end
                end
                save Temp\shipments.mat shipments
                if isfile([info.destinationfolder_sorted '\' info.filename]) == 0
                    createshipmentoverview('RemoteShipments_FB',info.year,info.month,info.day);
                end
            end
            sortshipmentexceltofolder(info);
            
        else
            % Do nothing with this file
            disp(['A destination could not be defined for ' info.filename]);
        end
        clear info
        clear shipments
    end
end

end