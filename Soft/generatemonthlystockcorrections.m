function generatemonthlystockcorrections()

disp('Calculating stock corrections - please wait');

load Temp\values.mat values
load Temp\onlinereport.mat onlinereport
load([values.salesordercheck_folder 'salesordercheck_v3.mat'],'salesordercheck_v3');

% Headers Online Report
headersOR = onlinereport(1,:);
col_CaseCode = catchcolumnindex({'CaseCode'},headersOR,1);
col_CaseCode = cell2mat(col_CaseCode(2,1));
col_ShippedYear = catchcolumnindex({'ShippedYear'},headersOR,1);
col_ShippedYear = cell2mat(col_ShippedYear(2,1));
col_ShippedMonth = catchcolumnindex({'ShippedMonth'},headersOR,1);
col_ShippedMonth = cell2mat(col_ShippedMonth(2,1));
col_ShippedDay = catchcolumnindex({'ShippedDay'},headersOR,1);
col_ShippedDay = cell2mat(col_ShippedDay(2,1));
col_HeelCupLeft = catchcolumnindex({'Heel_cup___left'},headersOR,1);
col_HeelCupLeft = cell2mat(col_HeelCupLeft(2,1));
col_HeelCupRight = catchcolumnindex({'Heel_cup___right'},headersOR,1);
col_HeelCupRight = cell2mat(col_HeelCupRight(2,1));
col_Shipped = catchcolumnindex({'Shipped'},headersOR,1);
col_Shipped = cell2mat(col_Shipped(2,1));
col_BuiltActor = catchcolumnindex({'BuiltActor'},headersOR,1);
col_BuiltActor = cell2mat(col_BuiltActor(2,1));
% Headers sales order check
col_caseid_soc = catchcolumnindex({'CaseID'},salesordercheck_v3.headers,1);
col_caseid_soc = cell2mat(col_caseid_soc(2,1));
col_itemnumber_soc = catchcolumnindex({'ItemNumber'},salesordercheck_v3.headers,1);
col_itemnumber_soc = cell2mat(col_itemnumber_soc(2,1));

% Define the headers of the item journal template
itemjournaltemplateheaders = {'Date','Adjustment','Documentnumber','Itemnumber','Description','Location','Project','Quantity'};
output = {};
output2 = {};
info.counter = 0;

% Define the date of Navision entry
if contains(['01' '03' '05' '07' '08' '10' '12'], values.RequestedMonth) == 1
    day_str = '31';
elseif contains(['04' '06' '09' '11'], values.RequestedMonth) == 1
    day_str = '30';
else
    day_str = '28';
end
info.output_date = [day_str '/' values.RequestedMonth '/' values.RequestedYear];

% Go over all rows
for cr = 2:size(onlinereport,1)
    % Check if the case is shipped
    if strcmp(char(onlinereport(cr,col_Shipped)),'-') == 0
        % Manipulate the month
        if cell2mat(onlinereport(cr,col_ShippedMonth)) < 10
            ActualMonth = ['0' num2str(cell2mat(onlinereport(cr,col_ShippedMonth)))];
        else
            ActualMonth = num2str(cell2mat(onlinereport(cr,col_ShippedMonth)));
        end
        % Check the shipment year/month
        if strcmp(values.RequestedYear,num2str(cell2mat(onlinereport(cr,col_ShippedYear)))) == 1 && strcmp(values.RequestedMonth,ActualMonth) == 1
            % Check if the case has a low heel cup on the left or right
            if (strcmp('Low',char(onlinereport(cr,col_HeelCupLeft))) == 1 || strcmp('Low',char(onlinereport(cr,col_HeelCupRight))) == 1)
                
                info.left = onlinereport(cr,col_HeelCupLeft);
                info.right = onlinereport(cr,col_HeelCupRight);
                
                info.builtactor = char(onlinereport(cr,col_BuiltActor));
                % Get the case ID
                info.caseid = onlinereport(cr,col_CaseCode);
                if strcmp(info.caseid,'RS20-FOC-VEC') == 1
                    test = 1;
                end
                % Get the position of the case ID in the salesordercheck
                row_caseid_soc = catchrowindex(info.caseid,salesordercheck_v3.data,col_caseid_soc);
                row_caseid_soc = cell2mat(row_caseid_soc(2,1));
                % Get the itemnumber
                itemnumber = char(salesordercheck_v3.data(row_caseid_soc,col_itemnumber_soc));
                info.itemnumber_full = itemnumber;
                
                % Is the order not from US and does it have a top layer?
                if strcmp('Flowbuilt Production',info.builtactor) == 0 && strcmp(itemnumber(15),'P') == 1
                    
                    % Get the top layer material
                    toplayer_itemnumber = itemnumber(4:14);
                    sheet_itemnumber = [toplayer_itemnumber 'S'];
                    sheet_itemnumber(5:9) = '99999';
                    % Get the additional top layer material
                    additional_itemnumber = itemnumber(17);
                    % Get the size
                    toplayer_size = itemnumber(10:11);
                    if strcmp(toplayer_size(1),'0')
                        toplayer_size = toplayer_size(2);
                    end
                    toplayer_size = str2double(toplayer_size);
                    % Get adults sizing
                    adultsizing = str2double(itemnumber(8));
                    % Shore value
                    shorevalue = itemnumber(13:14);
                    % Material
                    material = itemnumber(4:5);
                    % Thickness
                    thickness = str2double(itemnumber(6));
                    
                    % Only select combination with preformed top layers. Exclude all sheet materials
                    if adultsizing == 1
                        if (strcmp(material(1),'1') == 1 && strcmp(shorevalue,'40') == 1 && toplayer_size < 16) ...
                                || (strcmp(material(1),'1')  == 1 && toplayer_size < 14) ...
                                || (strcmp(material,'20') && toplayer_size > 3 && toplayer_size < 14) ...
                                || (strcmp(material,'30') && thickness == 3 && toplayer_size > 2 && toplayer_size < 16)
                            
                            % In this loop, the conversion will be performed;
                            info.counter = info.counter + 1;
                            
                            % If EVA
                            if strcmp(sheet_itemnumber(1),'1') == 1
                                % If 3 mm, adjust to 4 mm sheet
                                if strcmp(sheet_itemnumber(3),'3') == 1
                                    sheet_itemnumber(3) = '4';
                                end
                                % Check if it has synthetic leather
                                if strcmp(sheet_itemnumber(2),'1') == 1
                                    % Overwrite the SL
                                    sheet_itemnumber(2) = '0';
                                    % Book a negative adjustment on X02 (as it needs to be added to the sheet material)
                                    [output] = generatemonthlystockcorrections_addrow(output,info,values,'X02','Negative Adjmt.',2);
                                end
                                % Correct for the 2mm sheet to be 40 shore instead of 20 shore (even though it is not 20 shore..)
                                if strcmp(sheet_itemnumber,'10209999920S') == 1
                                    sheet_itemnumber = '10209999940S';
                                end
                                % Book a negative adjustment on the EVA sheet material
                                [output] = generatemonthlystockcorrections_addrow(output,info,values,sheet_itemnumber,'Negative Adjmt.',2);
                            end
                            
                            % If PU soft
                            if strcmp(sheet_itemnumber(1:2),'20') == 1
                                % If 3 mm, adjust to 4 mm sheet
                                if strcmp(sheet_itemnumber(3),'3') == 1
                                    sheet_itemnumber(3) = '4';
                                end
                                [output] = generatemonthlystockcorrections_addrow(output,info,values,sheet_itemnumber,'Negative Adjmt.',2);
                                %                         % Check if it has synthetic leather
                                %                         if strcmp(sheet_itemnumber(2),'0') == 1
                                %                             % Book a negative adjustment on PU soft black fabric
                                %                             [output] = generatemonthlystockcorrections_addrow(output,info,values,sheet_itemnumber,'Negative Adjmt.',2);
                                %                         elseif strcmp(sheet_itemnumber(2),'1') == 1
                                %                             % Book a negative adjustment on PU soft synthetich leather
                                %                             % !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                                %                         end
                            end
                            
                            % IF EVA carbon
                            if strcmp(sheet_itemnumber(1),'3') == 1
%                                 % If 3 mm, adjust to 4 mm sheet
%                                 if strcmp(sheet_itemnumber(3),'3') == 1
%                                     sheet_itemnumber(3) = '4';
%                                 end
                                [output] = generatemonthlystockcorrections_addrow(output,info,values,sheet_itemnumber,'Negative Adjmt.',2);
                            end
                            
                            % Check if it has an additional top layer
                            if strcmp(additional_itemnumber,'1') == 1
                                % Book a negative adjustment on X01 (as it needs to be added to the sheet material)
                                [output] = generatemonthlystockcorrections_addrow(output,info,values,'X01','Negative Adjmt.',2);
                            elseif strcmp(additional_itemnumber,'2') == 1
                                % Book a negative adjustment on X02 (as it needs to be added to the sheet material)
                                [output] = generatemonthlystockcorrections_addrow(output,info,values,'X02','Negative Adjmt.',2);
                            elseif strcmp(additional_itemnumber,'3') == 1
                                % Book a negative adjustment on X03 (as it needs to be added to the sheet material)
                                [output] = generatemonthlystockcorrections_addrow(output,info,values,'X03','Negative Adjmt.',2);
                            end
                            
                            % Do a positive adjustment on the required top layer materials
                            [output] = generatemonthlystockcorrections_addrow(output,info,values,[toplayer_itemnumber 'L'],'Positive Adjmt.',1);
                            [output] = generatemonthlystockcorrections_addrow(output,info,values,[toplayer_itemnumber 'R'],'Positive Adjmt.',1);
                        end
                    end
                end
            end
        end
    end
end

% Write the data to file
output2(1,:) = itemjournaltemplateheaders;
output2(2:size(output,1)+1,:) = output(:,8:15);
xlswrite([values.currentversion 'Output\Navision\20' values.y values.mo values.d '_CorrectionLowheelInsoles_ItemJournalTemplate.xlsx' ],output2);

end