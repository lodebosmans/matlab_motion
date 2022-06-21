function printshipmentdocuments()

load Temp\values.mat values
load Temp\list_shipments.mat list_shipments
load Temp\customers.mat customers

nroffields = size(list_shipments,1);
% Make the overview of the shipments
for x = 1:nroffields
    fieldname = char(list_shipments(x,1));
    % Get the index of the underscores
    k = strfind(fieldname,'_');
    list_shipments(x,2) = {fieldname(1,k(1,3)+1:k(1,4)-1)};
    % Load the shipment file
    load(['Temp\shipment_' char(list_shipments(x,2)) '.mat'],'Shipment');
    eval(['Shipment_' char(list_shipments(x,2)) ' = Shipment;']);
    eval(['list_shipments(x,3) = {Shipment_' char(list_shipments(x,2)) '.ShipmentRowNumber};']);
    clear Shipment
    clear k
end
clear x
list_shipments(1:end,4) = {0};
col_shipmentnumber = 2;
col_excelrownumber = 3;
col_printed = 4;

goon = 1;

while goon == 1
    
    tobeprinted = {};
    printed = {};
    tobeprinted_size = 0;
    printed_size = 0;
    print_worldship = 0;
    print_deliverynote = 0;
    
    % Make the overview of the shipments
    for x = 1:nroffields
        % Get the row number
        shipmentnumber = char(list_shipments(x,2));
        % Get the ship to number
        customernumber = eval(['Shipment_' shipmentnumber '.ShipTo.CustomerNumber_']);
        
        if size(customernumber,2) == 5
            customernumber = [customernumber '        '];
        end
        label = [shipmentnumber ' - ' customernumber ' - ' eval(['Shipment_' shipmentnumber '.ShipTo.CompanyOrName'])];
        tobeprinted_size = size(tobeprinted,1);
        printed_size = size(printed,1);
        if cell2mat(list_shipments(x,col_printed)) == 0
            tobeprinted(tobeprinted_size+1,:) = {label};
        else
            printed(printed_size+1,:) = {label};
        end
        
    end
    clear x
    tobeprinted_size = size(tobeprinted,1);
    printed_size = size(printed,1);
    
    % Sort the variables (if necessary)
    tobeprinted = sortrows(tobeprinted);
    printed = sortrows(printed);
    
    % Get all data together
    lines = '-----------------';
    notyetprinted = 'Not yet printed';
    alreadyprinted = 'Already printed';
    overview = {notyetprinted;lines};
    overview(end+1:end+tobeprinted_size,1) = tobeprinted;
    overview(end+1:end+3,:) = {' ';alreadyprinted;lines};
    overview(end+1:end+printed_size,1) = printed;
    
    % Request which shipment should be printed
    [indx,tf] = listdlg('PromptString','Select the shipment you would like to print','SelectionMode','single','ListString',overview,'ListSize',[500 600],'OKString','Select','CancelString','Exit shipments');
    
    if tf == 1 && strcmp(char(overview(indx,1)),lines) == 0  && strcmp(char(overview(indx,1)),notyetprinted) == 0 ...
            && strcmp(char(overview(indx,1)),alreadyprinted) == 0 && strcmp(char(overview(indx,1)),' ') == 0
        
        
        
        % ---------------------
        % For regular shipments
        % ---------------------
        
        % Get the shiptonr
        k = strfind(char(overview(indx,1)),'-');
        temp = char(overview(indx,1));
        selectedrow =  temp(1,1:k(1,1)-2);
        
        
        selectedoption_nr = find(strcmp(list_shipments(:,col_shipmentnumber),selectedrow ));
        selectedoption = num2str(cell2mat(list_shipments(selectedoption_nr,col_excelrownumber)));
        
        fieldname2_shipmentnumber = ['Shipment_' selectedrow];
        fieldname2_excelrownumber = ['Shipment_' num2str(selectedoption)];
        fieldname2_deliverynote = ['DeliveryNote_' selectedoption];
        
        % Print that shizzle
        answer = questdlg('What would you like to print?', ...
            'Please select an option', ...
            'WorldShip XML + delivery note','WorldShip XML','Delivery note','WorldShip XML + delivery note');
        % Handle response
        switch answer
            case 'WorldShip XML + delivery note'
                print_worldship = 1;
                print_deliverynote = 1;
            case 'WorldShip XML'
                print_worldship = 1;
            case 'Delivery note'
                print_deliverynote = 1;
        end
        
        if print_worldship == 1
            % filename
            % copyfile([values.currentversion '\Output\WorldShip\*Shipment*' shipments.(fieldname2).deliveryname '.xml'],[values.worldshipxmlfilefolder currentfilename '.xml']);
            copyfile([pwd '\Output\WorldShip\*' fieldname2_excelrownumber '*.xml'],[values.worldshipxmlfilefolder]);
        end
        if print_deliverynote == 1
            if values.MenuSelectPrinter > 1
                % Find file in folder
                listing = dir([pwd '\Output\DeliveryNotes']);
                
                for y = 1:size(listing,1)
                    if contains(listing(y).name,fieldname2_deliverynote) == 1 && strcmp(listing(y).name(end-3:end),'xlsx') == 1
                        target = y;
                    end
                end
                
                % Get the amount of pages
                nrofpages = ceil(eval(['size(' fieldname2_shipmentnumber '.CaseIDs,1)'])/25);
                if nrofpages == 0
                    nrofpages = 1;
                end
                
                file = [pwd '\Output\DeliveryNotes\' listing(target).name];
                % Print it
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
                Excel.ActiveWorkbook.PrintOut(1,nrofpages,values.NrOfCopies,'False',values.MenuSelectPrinter_str); % second is nr of pages
                invoke(Excel,'Quit');
                delete(Excel);
                clear Excel;
            end
        end
        
        % Indicate it has been printed in the shipments variable.
        list_shipments(selectedoption_nr,col_printed) = {1};
        
        % --------------------------------
        % For shipments with a subdelivery
        % --------------------------------
        
        % If there are subdeliveries, print these as well
        temp_fieldnames = fields(eval(fieldname2_shipmentnumber));
        
        
        if sum(contains(temp_fieldnames,{'Subdeliveries'})) == 1
            % Get the amount of subdeliveries
            subdelivery_fieldnames = fields(eval([fieldname2_shipmentnumber '.Subdeliveries']));
            subdelivery_nroffields = size(subdelivery_fieldnames,1);
            for i = 1:subdelivery_nroffields
                %                 % Get the shiptonr
                               subdelivery_selected = char(subdelivery_fieldnames(i,1));
                %                 fieldname3 = strrep(subdelivery_selected,' ','');
                %                 % Get the customer name
                %                 customernr_c = strrep(fieldname3,'D','C');
                %                 companyname = customers.(customernr_c).Company;
                
                test = num2str(eval([fieldname2_shipmentnumber '.Subdeliveries.' subdelivery_selected '.ShipmentRowNumber;']));
                fieldname3_deliverynote = ['DeliveryNote_' test];
                
                if print_deliverynote == 1
                    if values.MenuSelectPrinter > 1
                        
                        for y = 1:size(listing,1)
                            if contains(listing(y).name,fieldname3_deliverynote) == 1 && strcmp(listing(y).name(end-3:end),'xlsx') == 1
                                target = y;
                            end
                        end
                        
                        % Get the amount of pages
                        nrofpages = ceil(eval(['size(' fieldname2_shipmentnumber '.Subdeliveries.' num2str(subdelivery_selected) '.CaseIDs,1)'])/25);
                        
                        file = [pwd '\Output\DeliveryNotes\' listing(target).name];
                        % Print it
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
                        Excel.ActiveWorkbook.PrintOut(1,nrofpages,values.NrOfCopies,'False',values.MenuSelectPrinter_str); % second is nr of pages
                        invoke(Excel,'Quit');
                        delete(Excel);
                        clear Excel;
                    end
                end
            end
        end
    elseif strcmp(char(overview(indx,1)),lines) == 1 || strcmp(char(overview(indx,1)),notyetprinted) == 1 ...
            || strcmp(char(overview(indx,1)),alreadyprinted) == 1 || strcmp(char(overview(indx,1)),' ') == 1
        goon = 1;
    else
        goon = 0;
    end
end


end




