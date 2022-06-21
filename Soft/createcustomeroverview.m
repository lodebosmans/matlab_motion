function createcustomeroverview()

load Temp\values.mat values
load Temp\UPSfile_customer.mat UPSfile_customer
load Temp\UPSfile_customagreements.mat UPSfile_customagreements

disp('Creating customer overview - please wait')

% First read the customer agreements
% ----------------------------------

% Read the headers
heardersagr = UPSfile_customagreements(2,:);
nrofcols = size(heardersagr,2);
for x = 1:nrofcols
    ca = char(UPSfile_customagreements(2,x));
    eval(['colagr.' ca ' = catchcolumnindex({''' ca '''},UPSfile_customagreements,2);']);
    eval(['colagr.' ca ' = cell2mat(colagr.' ca '(2,1));']);
end
clear ca
clear x

% Find the row of the last agreement
lastrowagreement = 4;
counter = 0;
while lastrowagreement == 4
    counter = counter + 1;
    if strcmp(UPSfile_customagreements(lastrowagreement+counter,colagr.CustomerNumber),'-') == 1
        lastrowagreement = lastrowagreement + counter-1;
    end
end

% Read all the agreements and sort per customer
customers = struct;
agreements = struct;
col_deliveryonly = catchcolumnindex({'DeliveryOnly'},UPSfile_customagreements,2);
col_deliveryonly = cell2mat(col_deliveryonly(2,1));

for ca = 4:lastrowagreement
    % Get the customer
    cnr = char(UPSfile_customagreements(ca,colagr.CustomerNumber));
    cnr = strrep(cnr,'-','_');
    % Display message
    disp(['Reading customer agreement ' num2str(ca) ' - ' cnr ' - please wait']);
    % Get the full line
    cafull = UPSfile_customagreements(ca,:);
    % Check if the customer of this agreement is already known.
    if isfield(agreements, cnr) == 0
        % If not, create the first entry for this customer.
        [agreements] = addagreement(agreements,cafull,heardersagr,cnr,1);
        % Add DeliveryOnly parameter to the customers variable.
        customers.(cnr).DeliveryOnly = agreements.(cnr).PriceAgreements.Agreement1.DeliveryOnly;
        % Add the alternative customer number to the customers variable
        customers.(cnr).AlternativeCustomerNumber = agreements.(cnr).PriceAgreements.Agreement1.AlternativeCustomerNumber;
    else
        % Count the number of present fields and add the new one.
        camount = eval(['length(fieldnames(agreements.' cnr '.PriceAgreements));']);
        [agreements] = addagreement(agreements,cafull,heardersagr,cnr,camount+1); 
        % Add DeliveryOnly parameter to the customers variable.
        customers.(cnr).DeliveryOnly = agreements.(cnr).PriceAgreements.(['Agreement' num2str(camount+1)]).DeliveryOnly;
    end    
end
clear cnr

% Secondly read the customer data and add the agreements
% ------------------------------------------------------
customers.overview = cell(0);
rowoverview = 1;

col.customerlabel = UPSfile_customer(2,:);
col_number = catchcolumnindex({'CustomerNumber'},col.customerlabel,1);
col_number = cell2mat(col_number(2,1));
col_name = catchcolumnindex({'Company'},col.customerlabel,1);
col_name = cell2mat(col_name(2,1));
col_obsolete = catchcolumnindex({'Obsolete'},col.customerlabel,1);
col_obsolete = cell2mat(col_obsolete(2,1));
% col_deliveryonly = catchcolumnindex({'DeliveryOnly'},col.customerlabel,1);
% col_deliveryonly = cell2mat(col_deliveryonly(2,1));

% Create some files for the customer databases
info_distributor_list = {'superfeet','go4d','go4d'; ...
                          'C0032','C0022','C0022-003'};
col_NotificationEmail = catchcolumnindex({'NotificationEmail'},UPSfile_customer(2,:),1);
col_NotificationEmail = cell2mat(col_NotificationEmail(2,1));
for z = 1:size(info_distributor_list,2)
%     temp = char(info_distributor_list(1,z));
    temp2 = char(info_distributor_list(2,z));
    eval(['customers_' strrep(temp2,'-','_') ' = cell(0);']);
    eval(['customers_' strrep(temp2,'-','_') '(1,:) = UPSfile_customer(2,1:col_NotificationEmail);']);
    eval(['customers_' strrep(temp2,'-','_') '(1,end+1) = {''AlternativeCustomerNumber''}; ']);
end


% Create the orders struct for if prices need to be calculated.
ordersstruct = struct;
if values.NavSalesOrderInput == 1
    for y = 2018:str2double(['20' values.y])
        ordersstruct.(['y' num2str(y)]).overview = cell(0);
        ordersstruct.(['y' num2str(y)]).RegSemi = 0;
        ordersstruct.(['y' num2str(y)]).RegFull = 0;
        ordersstruct.(['y' num2str(y)]).SlimSemi = 0;
        ordersstruct.(['y' num2str(y)]).SlimFull = 0;
        for mo = 1:12
            ordersstruct.(['y' num2str(y)]).(['m' num2str(mo)]) = cell(0);
        end
    end
end

nrofrows = size(UPSfile_customer,1);
for cr = 3:nrofrows
    %     disp(cr);
    %     if cr == 466
    %         stopt = 1;
    %     end
    if isempty(UPSfile_customer(cr,col_number)) == 0 && length(char(UPSfile_customer(cr,col_name))) > 0 % && strcmp(UPSfile_customer(cr,col_obsolete),'Yes') == 0
        cnr = char(UPSfile_customer(cr,col_number));
        disp(['Processing customer ' cnr]);
        % Skip all the marketing numbers
        if strcmp(cnr(1,1:5),'CXXXX') == 0
        %if strcmp(cnr(1,1:5),'C0002') == 0
            %disp(cnr)
            cnr = strrep(cnr,'-','_'); % In the path, always keep the full company number
%             % If only a delivery address, remove the suffix of the customer number (but not in the delivery number)
            %delonly = upper(char(UPSfile_customer(cr,col_deliveryonly)));
            if isfield(customers, cnr) == 1
                delonly = upper(customers.(cnr).DeliveryOnly);
            else
                delonly = '-';
            end
            if strcmp(delonly,'YES') == 1
                tmp = char(UPSfile_customer(cr,col_number));
                customers.overview(rowoverview,1) = {tmp}; % In the overview, keep the suffix.
                customers.(cnr).CustomerNumber = tmp(1:5); % In the company struct, delete the suffix.
            else
                customers.(cnr).CustomerNumber = char(UPSfile_customer(cr,col_number));
                customers.overview(rowoverview,1) = {customers.(cnr).CustomerNumber};
            end
            
            
            
            % General info
            % --------------------------
            temp = catchcolumnindex({'Company'},col.customerlabel,1);
            customers.(cnr).Company = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %customers.overview(rowoverview,3) = {customers.(cnr).Company};
            %
            temp = catchcolumnindex({'ContactPerson'},col.customerlabel,1);
            customers.(cnr).ContactPerson = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'Email'},col.customerlabel,1);
            customers.(cnr).Email = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'VAT/EIN'},col.customerlabel,1);
            if isstring(UPSfile_customer(cr,cell2mat(temp(2,1)))) == 1
                customers.(cnr).VatEin = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            else
                customers.(cnr).VatEin = num2str(cell2mat(UPSfile_customer(cr,cell2mat(temp(2,1)))));
            end
            %
            temp = catchcolumnindex({'Phone'},col.customerlabel,1);
            if isstring(UPSfile_customer(cr,cell2mat(temp(2,1)))) == 1
                customers.(cnr).Phone = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            else
                customers.(cnr).Phone = num2str(cell2mat(UPSfile_customer(cr,cell2mat(temp(2,1)))));
            end
            
            
            % Clinic address
            % --------------------------
            temp = catchcolumnindex({'ClinicContactPerson'},col.customerlabel,1);
            customers.(cnr).ClinicAddress.ClinicContactPerson = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'ClinicAddress1'},col.customerlabel,1);
            customers.(cnr).ClinicAddress.ClinicAddress1 = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'ClinicAddress2'},col.customerlabel,1);
            customers.(cnr).ClinicAddress.ClinicAddress2 = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'ClinicAddress3'},col.customerlabel,1);
            customers.(cnr).ClinicAddress.ClinicAddress3 = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'ClinicPostalCode'},col.customerlabel,1);
            if isstring(UPSfile_customer(cr,cell2mat(temp(2,1)))) == 1
                customers.(cnr).ClinicAddress.ClinicPostalCode = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            else
                customers.(cnr).ClinicAddress.ClinicPostalCode = num2str(cell2mat(UPSfile_customer(cr,cell2mat(temp(2,1)))));
            end
            %
            temp = catchcolumnindex({'ClinicCity'},col.customerlabel,1);
            customers.(cnr).ClinicAddress.ClinicCity = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'ClinicState'},col.customerlabel,1);
            customers.(cnr).ClinicAddress.ClinicState = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'ClinicCountry'},col.customerlabel,1);
            customers.(cnr).ClinicAddress.ClinicCountry = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'ClinicPhone'},col.customerlabel,1);
            if isstring(UPSfile_customer(cr,cell2mat(temp(2,1)))) == 1
                customers.(cnr).ClinicAddress.ClinicPhone = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            else
                customers.(cnr).ClinicAddress.ClinicPhone = num2str(cell2mat(UPSfile_customer(cr,cell2mat(temp(2,1)))));
            end
            
            
            % Invoice address
            % --------------------------
            temp = catchcolumnindex({'InvoiceContactPerson'},col.customerlabel,1);
            customers.(cnr).InvoiceAddress.InvoiceContactPerson = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'InvoiceAddress1'},col.customerlabel,1);
            customers.(cnr).InvoiceAddress.InvoiceAddress1 = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'InvoiceAddress2'},col.customerlabel,1);
            customers.(cnr).InvoiceAddress.InvoiceAddress2 = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'InvoiceAddress3'},col.customerlabel,1);
            customers.(cnr).InvoiceAddress.InvoiceAddress3 = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'InvoicePostalCode'},col.customerlabel,1);
            if isstring(UPSfile_customer(cr,cell2mat(temp(2,1)))) == 1
                customers.(cnr).InvoiceAddress.InvoicePostalCode = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            else
                customers.(cnr).InvoiceAddress.InvoicePostalCode = num2str(cell2mat(UPSfile_customer(cr,cell2mat(temp(2,1)))));
            end
            %
            temp = catchcolumnindex({'InvoiceCity'},col.customerlabel,1);
            customers.(cnr).InvoiceAddress.InvoiceCity = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'InvoiceState'},col.customerlabel,1);
            customers.(cnr).InvoiceAddress.InvoiceState = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'InvoiceCountry'},col.customerlabel,1);
            customers.(cnr).InvoiceAddress.InvoiceCountry = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'InvoicePhone'},col.customerlabel,1);
            if isstring(UPSfile_customer(cr,cell2mat(temp(2,1)))) == 1
                customers.(cnr).InvoiceAddress.InvoicePhone = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            else
                customers.(cnr).InvoiceAddress.InvoicePhone = num2str(cell2mat(UPSfile_customer(cr,cell2mat(temp(2,1)))));
            end
            %
%             temp = catchcolumnindex({'InvoiceTo'},col.customerlabel,1);
%             customers.(cnr).InvoiceAddress.InvoiceTo = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
%             
            
            % Delivery addresss
            % --------------------------
            temp = catchcolumnindex({'DeliveryFacilityName'},col.customerlabel,1);
            customers.(cnr).DeliveryAddress.DeliveryFacilityName = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %customers.overview(rowoverview,3) = {customers.(cnr).DeliveryAddress.DeliveryFacilityName};
            %
            temp = catchcolumnindex({'DeliveryContactPerson'},col.customerlabel,1);
            customers.(cnr).DeliveryAddress.DeliveryContactPerson = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'DeliveryAddress1'},col.customerlabel,1);
            customers.(cnr).DeliveryAddress.DeliveryAddress1 = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'DeliveryAddress2'},col.customerlabel,1);
            customers.(cnr).DeliveryAddress.DeliveryAddress2 = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'DeliveryAddress3'},col.customerlabel,1);
            customers.(cnr).DeliveryAddress.DeliveryAddress3 = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'DeliveryPostalCode'},col.customerlabel,1);
            if isstring(UPSfile_customer(cr,cell2mat(temp(2,1)))) == 1
                customers.(cnr).DeliveryAddress.DeliveryPostalCode = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            else
                customers.(cnr).DeliveryAddress.DeliveryPostalCode = num2str(cell2mat(UPSfile_customer(cr,cell2mat(temp(2,1)))));
            end
            %
            temp = catchcolumnindex({'DeliveryCity'},col.customerlabel,1);
            customers.(cnr).DeliveryAddress.DeliveryCity = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'DeliveryState'},col.customerlabel,1);
            customers.(cnr).DeliveryAddress.DeliveryState = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'DeliveryCountry'},col.customerlabel,1);
            customers.(cnr).DeliveryAddress.DeliveryCountry = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'DeliveryPhone'},col.customerlabel,1);
            if isstring(UPSfile_customer(cr,cell2mat(temp(2,1)))) == 1
                customers.(cnr).DeliveryAddress.DeliveryPhone = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            else
                customers.(cnr).DeliveryAddress.DeliveryPhone = num2str(cell2mat(UPSfile_customer(cr,cell2mat(temp(2,1)))));
            end
            %
            temp = catchcolumnindex({'CountryCode'},col.customerlabel,1);
            customers.(cnr).DeliveryAddress.DeliveryCountryCode = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'StateCode'},col.customerlabel,1);
            customers.(cnr).DeliveryAddress.DeliveryStateCode = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'EU'},col.customerlabel,1);
            customers.(cnr).DeliveryAddress.DeliveryEU = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
%             temp = catchcolumnindex({'ShipTo'},col.customerlabel,1);
%             customers.(cnr).DeliveryAddress.SendTo = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
%             
            % OMS details
            % -----------
            temp = catchcolumnindex({'SurgeonOMS'},col.customerlabel,1);
            customers.(cnr).OMS.SurgeonOMS = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'CompanyNameOMS'},col.customerlabel,1);
            customers.(cnr).OMS.CompanyNameOMS = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            customers.overview(rowoverview,2) = {customers.(cnr).OMS.CompanyNameOMS};
            %
            temp = catchcolumnindex({'DeliveryOffice'},col.customerlabel,1);
            customers.(cnr).OMS.DeliveryOffice = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            customers.overview(rowoverview,3) = {customers.(cnr).OMS.DeliveryOffice};
            %
            temp = catchcolumnindex({'Obsolete'},col.customerlabel,1);
            customers.(cnr).OMS.Obsolete = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            if strcmp(customers.(cnr).OMS.Obsolete,'Yes') == 1
                customers.overview(rowoverview,4) = cellstr('Obsolete');
            else
                customers.overview(rowoverview,4) = cellstr('Active');
            end
            
            % Emails
            % ------
            temp = catchcolumnindex({'InvoiceEmail'},col.customerlabel,1);
            customers.(cnr).InvoiceEmail = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            %
            temp = catchcolumnindex({'NotificationEmail'},col.customerlabel,1);
            customers.(cnr).NotificationEmail = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            
            % Extra
            % -----
            temp = catchcolumnindex({'MarketSegment'},col.customerlabel,1);
            customers.(cnr).MarketSegment = char(UPSfile_customer(cr,cell2mat(temp(2,1))));
            
            
            
            % Add the price agreements, if they are available.            
            if isfield(agreements,cnr)
                customers.(cnr).PriceAgreements = agreements.(cnr).PriceAgreements;
                %
                [agrnr] = getcurrentpriceagreement(customers,cnr,2000+str2double(values.y),str2double(values.mo),str2double(values.d));
                
                if agrnr ~= 0 % This part should be rewritten to get this info straight from the agreement instead of from this copied information.
                    % Invoicing and pricing information.
                    % ----------------------------------
                    info = {'DeliveryOnly','ShipTo','Subdelivery','InvoiceTo','Agent','Importer','UPSAccountNumber','Main','PricingMethod','Currency','StandardPrice','SplitTypeCounter', ...
                        'Discount','NoTopLayerSN','NoAddressLabel','ActionMessage','PrintCommercialInvoice','FixedShipmentDay','NeutralShipmentBox','TransportFee'};
                    sizeinfo = size(info,2);
                    for ci = 1: sizeinfo
                        ci_str = char(info(1,ci));
                        eval(['customers.' cnr '.' ci_str ' = customers.' cnr '.PriceAgreements.Agreement' num2str(agrnr) '.' ci_str ';']);
                    end
                end
            end
            
            

            rowoverview = rowoverview + 1;
        end
        % Compensate for the fact that marketing orders do not have agreements.
        if strcmp(cnr(1,1:5),'C0002') == 1
            info = {'DeliveryOnly','Subdelivery','InvoiceTo','Agent','Main','PricingMethod','Currency','StandardPrice','SplitTypeCounter', ...
                'Discount','NoTopLayerSN','NoAddressLabel'};
            sizeinfo = size(info,2);
            for ci = 1: sizeinfo
                ci_str = char(info(1,ci));
                eval(['customers.' cnr '.' ci_str ' = ''-'';']);
            end
            cstn_d = cnr;
            cstn_d(1) = 'D';
            strrep(cstn_d,'_','-');            
            eval(['customers.' cnr '.ShipTo = cstn_d;']);
        end
        
        % If necessary, add this customer to the customeroverview of its agent
        if values.NavSalesOrderInput == 1 || values.UpdateOnlineReport == 1
            if isfield(customers.(cnr),'Agent') == 1
                if any(strcmp(info_distributor_list(2,:),customers.(cnr).Agent)) == 1
                    % Filter out C1078. A special customer.
                    if (strcmp(cnr,'C1078') == 1 && strcmp(char(customers.overview(end,4)),'Obsolete') == 0 ) || strcmp(cnr,'C1078') == 0
                        eval(['customers_' strrep(customers.(cnr).Agent,'-','_') '(end+1,1:end-1) = UPSfile_customer(cr,1:col_NotificationEmail);' ]);
                        eval(['customers_' strrep(customers.(cnr).Agent,'-','_') '(end,end) = {customers.' cnr '.AlternativeCustomerNumber};' ]);
                    end
                end
            end
        end
    end
end

if values.NavSalesOrderInput == 1 || values.UpdateOnlineReport == 1
    for q = 1:size(info_distributor_list,2)
        %     cds_name = char(info_distributor_list(2,q));
        cds_name = char(info_distributor_list(1,q));
        cds_name_code = strrep(char(info_distributor_list(2,q)),'-','_');
        filename = [values.onlinereportfolder '20' values.y values.mo values.d '_' values.h values.mi values.s '_' cds_name_code '_' cds_name '_customers.csv'];    %must end in csv
        writetable(cell2table(eval(['customers_' cds_name_code])), filename, 'writevariablenames', false, 'quotestrings', true);
        if strcmp(values.version,'liveversion') == 1
            versionvariable = '';
        else
            versionvariable = '-test';
        end
        % Activate the python script to send csv file to the footscan portal.
        command = ['C:\Users\mathlab\AppData\Local\Programs\Python\Python310\python.exe ' values.currentversion 'Soft\update-rsprint-customers' versionvariable '.py ' cds_name ' ' filename];
        system(command);
    end
end



save Temp\customers.mat customers

end