clear all
close all
clc

% Read the cases to be linked
filename = 'C:\Users\mathlab\Documents\Matlab\201900521_Interface_TestEnvironment_DoNotUse\Input\Orthocare Solutions.xlsx';
[num,text,OS_cases] = xlsread(filename);
nrofcases = size(OS_cases,1);

% read the shipment overview
filename = 'C:\Users\mathlab\Documents\Matlab\201900521_Interface_TestEnvironment_DoNotUse\Input\20160000_OverviewShipments_v2.xlsx';
[num,text,shipments] = xlsread(filename,2);
nrofshipments = size(shipments,1);
shipments(cellfun(@(x) isnumeric(x) && isnan(x), shipments)) = {'-'};

col_cusnr = catchcolumnindex({'CustomerNumber'},shipments,1);
col_cusnr = cell2mat(col_cusnr(2,1));
col_serial = catchcolumnindex({'SerialNumbers'},shipments,1);
col_serial = cell2mat(col_serial(2,1));
col_track = catchcolumnindex({'Trackingnumber'},shipments,1);
col_track = cell2mat(col_track(2,1));
col_year = catchcolumnindex({'Year'},shipments,1);
col_year = cell2mat(col_year(2,1));
col_month = catchcolumnindex({'Month'},shipments,1);
col_month = cell2mat(col_month(2,1));
col_day = catchcolumnindex({'Day'},shipments,1);
col_day = cell2mat(col_day(2,1));

counter = 0;
OScounter = 0;

for cs = 1:nrofshipments
    if cs ~= 1426 && cs ~= 2192 
        temp = shipments(cs,col_cusnr);
        if iscellstr(temp) == 1
            temp = char(temp);
        else
            temp = num2str(cell2mat(temp));
        end
        %     if strcmp(temp,'84.1') == 1 || strcmp(temp,'87.1') == 1
        OScounter = OScounter + 1;
        disp(['Amount of shipments to Orthocare solutions/training room = ' num2str(OScounter) ]);
        
        % Get the content of the shipment
        content = shipments(cs,col_serial);
        if iscellstr(content) == 1
            content = char(content);
        else
            content = num2str(cell2mat(content));
        end
        % Get the case IDs
        [output] = readmanualcaseids({},'None',0,content,1);
        nrofcasesinshipment = size(output,1);
        if nrofcasesinshipment > 0
            % Get the other data
            date_y = shipments(cs,col_year);
            if iscellstr(date_y) == 1
                date_y = char(date_y);
            else
                date_y = num2str(cell2mat(date_y));
            end
            date_m = shipments(cs,col_month);
            if iscellstr(date_m) == 1
                date_m = char(date_m);
            else
                date_m = num2str(cell2mat(date_m));
            end
            date_d = shipments(cs,col_day);
            if iscellstr(date_d) == 1
                date_d = char(date_d);
            else
                date_d = num2str(cell2mat(date_d));
            end
            %         date_y = char(shipments(cs,col_year));
            %         date_m = char(shipments(cs,col_month));
            %         date_d = char(shipments(cs,col_day));
            trackingnumber = char(shipments(cs,col_track));
            % Check if they are in the other file
            for x = 1:nrofcasesinshipment
                ccid = output(x,1);
                
                if strcmp('RS16-EFA-OKO',char(ccid)) == 1
                    stopnow = 1;
                end
                try
                    temp = catchrowindex(ccid,OS_cases,2);
                    temp = cell2mat(temp(2,1));
                    counter = counter + 1;
                    disp(['Amount of cases found = ' num2str(counter) ]);
                    
                    OS_cases(temp,3) = {trackingnumber};
                    OS_cases(temp,4) = {date_y};
                    OS_cases(temp,5) = {date_m};
                    OS_cases(temp,6) = {date_d};
                    
                    
                    
                catch
                    
                end
            end
        end
    end
    
    %
    %     end
end

xlswrite('Output\WalterReed.xlsx',OS_cases);

disp('Staaaph');