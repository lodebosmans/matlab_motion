function createinvoices()

load Temp\values.mat values
load Temp\onlinereport.mat onlinereport
load Temp\UPSfile_customer.mat UPSfile_customer

% Get the desired year and month - convert to numbers
shipped.year = str2double(values.MenuYear_str);
shipped.month = values.MenuMonth - 1;

stringstofetch_invoice = {'CaseCode','NameMerged','ShippedDateEU','DeliveryOfficeName','InsoleType','Top material','ReferenceID','HospitalName','ServiceType','Surgeon',...
                          'CaseCanceled','CaseId','ShippedYear','ShippedMonth','ShippedDay','CustomerFirstName','CustomerSurname'};
[stringstofetch_invoice] = catchcolumnindex(stringstofetch_invoice,onlinereport,1);

% Get all cases in this desired year and month
if values.MenuYear + values.MenuMonth >= 4 
    disp(['Creating invoices - please wait - gathering data for ' values.MenuMonth_str ' ' values.MenuYear_str])
    casecounter = 0;
    % Get the raw data for a month
    nrofcases_OR = size(onlinereport,1);
    for currentcase = 2:nrofcases_OR
        if strcmp(char(onlinereport(currentcase,cell2mat(stringstofetch_invoice(2,3)))),'//') == 0 % first check on case cancelled
            if cell2mat(onlinereport(currentcase,cell2mat(stringstofetch_invoice(2,13)))) == shipped.year
                if cell2mat(onlinereport(currentcase,cell2mat(stringstofetch_invoice(2,14)))) == shipped.month
                    if isnan(cell2mat(onlinereport(currentcase,cell2mat(stringstofetch_invoice(2,11))))) % double check on case cancelled
                        casecounter = casecounter + 1;
                        invoice.raw(casecounter,:) = onlinereport(currentcase,:);
                    end
                end
            end
        end
    end
    clear currentcase
% Get the data for a custom list of case IDs
elseif isempty(values.ManualInvoice) == 0
    disp('Creating invoices - please wait - gathering data for manually inserted case IDs');
    caseids_invoice = cell(0,1);
    [caseids_invoice] = readmanualcaseids(values,caseids_invoice,'ManualInvoice');
    % Transpose to match format expected for next function
    caseids_invoice = caseids_invoice';
    % Find indexes of the cases
    [caseids_invoice] = catchrowindex(caseids_invoice,onlinereport,2);
    % Get the data for the case IDs
    for currentcase = 1:size(caseids_invoice,2)        
        invoice.raw(currentcase,:) = onlinereport(cell2mat(caseids_invoice(2,currentcase)),:);        
    end
else
    disp('No valid selection of invoice input');
end

% Sort by customer and creation date
invoice.sorted = sortrows(invoice.raw,[cell2mat(stringstofetch_invoice(2,8)),cell2mat(stringstofetch_invoice(2,11))]);

% Only keep what is necessary for the invoice addendum
for y = 1:7
    invoice.filtered(:,y) = invoice.sorted(:,cell2mat(stringstofetch_invoice(2,y)));
end
clear y

% Replace the top layer with only the info of the additional synthetic top layer (and the additional cost related to it)
invoice.filtered(cellfun(@(x) strcmp(x,'EVA Synthetic Leather'), invoice.filtered)) = {'Synthetic leather'};
invoice.filtered(cellfun(@(x) strcmp(x,'PU Soft Synthetic Leather'), invoice.filtered)) = {'Synthetic leather'};
invoice.filtered(cellfun(@(x) strcmp(x,'EVA'), invoice.filtered)) = {''};
invoice.filtered(cellfun(@(x) strcmp(x,'PU Soft'), invoice.filtered)) = {''};

% Filter out the NaNs
invoice.filtered(cellfun(@(x) isnumeric(x) && isnan(x), invoice.filtered)) = {''};

% % Copy the UPS file with all customer information
% copyfile('O:\StandardOperations\UPS\RS Print\20170731_OverviewShipments_v5_DoNotUse.xlsx','Temp\UPSfile_customer.xlsx');
% 
% % Open UPS file with all customers
% [data.num,data.txt,UPSfile_customer] = xlsread('Temp\UPSfile_customer.xlsx',1);

% Add the pricing data for each case
Stringstofetch_ups = {'CompanyNameOMS','Currency','TaxRate','StandardPrice','PriceRegular','PriceCycling','PriceGolf','InvoiceMonthly','InvoiceAlternatively','CountryCode'};
[Stringstofetch_ups] = catchcolumnindex(Stringstofetch_ups,UPSfile_customer,2);

% Add the company temporarily, so we match it with the company name in the OMS based on the UPS file
invoice.filtered(:,13) = invoice.sorted(:,cell2mat(stringstofetch_invoice(2,8)));

invoice.clientspecific = struct;
invoice.nrofcustomers = 0;
invoice.customerlist = cell(1,1);
for y = 1:size(invoice.filtered,1)
    % Get the current company (in a few formats)
    current.company = invoice.filtered(y,13);
    current.company_str = char(invoice.filtered(y,13));
    current.company_str2 = current.company_str(1:end-8);
    current.company_str3 = regexprep(current.company_str2,' ','_');
    current.company_str4 = ['client_' current.company_str3];
    %disp(current.company_str2)
    % Get the current insole type
    current.insoletype = char(invoice.filtered(y,5));

    % Locate the company in the UPS file
    [current.company] = catchrowindex(current.company,UPSfile_customer,cell2mat(Stringstofetch_ups(2,1)));
%     disp(y)
%     disp(current.company)
%     if y == 192
%         donothing = 1;
%     end
    current.destinationcountry = char(UPSfile_customer(cell2mat(current.company(2,1)),cell2mat(Stringstofetch_ups(2,10))));
    % Get the price and taxrate of the insole type
    if strcmp('Synthetic leather',char(invoice.filtered(y,6))) || strcmp(current.insoletype,'Golf')
        invoice.filtered(y,10) = UPSfile_customer(cell2mat(current.company(2,1)),cell2mat(Stringstofetch_ups(2,7)));
        % Except for some customers
        exceptions = {'Zorg en Farma','TZW Limburg','TZW West Vlaanderen','TZW Midden Vlaanderen','RSscan Lab Ltd','MetalGenTech Co Ltd','TIM fysiotherapie BV','Livit Orthopedie bv','On Feet','Fotspesialisten Klinikk Elin',...
                      'Fotterapeut Sissel Heiberg','Groep SAM','Sportpodo DHaese'};
        idx = strcmp(current.company_str2,exceptions(1,:));
        if sum(idx) > 0 || strcmp(current.destinationcountry,'US') == 1
            invoice.filtered(y,10) = UPSfile_customer(cell2mat(current.company(2,1)),cell2mat(Stringstofetch_ups(2,5)));
        end
    elseif strcmp(current.insoletype,'Comfort') || strcmp(current.insoletype,'Running') || strcmp(current.insoletype,'Soccer') || strcmp(current.insoletype,'Narrow')
        invoice.filtered(y,10) = UPSfile_customer(cell2mat(current.company(2,1)),cell2mat(Stringstofetch_ups(2,5)));
    elseif strcmp(current.insoletype,'Cycling') || strcmp(current.insoletype,'Alpine ski') || strcmp(current.insoletype,'Nordic ski')
        invoice.filtered(y,10) = UPSfile_customer(cell2mat(current.company(2,1)),cell2mat(Stringstofetch_ups(2,6)));
    %elseif strcmp(current.insoletype,'Golf')
    %    invoice.filtered(y,10) = UPSfile_customer(cell2mat(current.company(2,1)),cell2mat(Stringstofetch_ups(2,7)));
    else
        x = char(invoice.filtered(y,1));
        disp(['Insole type not known for ' x]);
        clear x
    end
    % Calculate the price
    invoice.filtered(y,11) = UPSfile_customer(cell2mat(current.company(2,1)),cell2mat(Stringstofetch_ups(2,3)));
    price = cell2mat(invoice.filtered(y,10)) * (1 + (cell2mat(invoice.filtered(y,11))/100));
    invoice.filtered(y,12) = num2cell(price);

    % Check if this client has already passed the revue
    if isfield(invoice.clientspecific,current.company_str4)
        % Add to the existing list of the customer
        [invoice] = addtospecificinvoice(current.company_str4,invoice,y);
    else
        % Create new customer in list
        invoice.nrofcustomers = invoice.nrofcustomers + 1;
        invoice.customerlist(invoice.nrofcustomers,1) = cellstr(current.company_str4);
        invoice.clientspecific.(current.company_str4).counter = 0;
        invoice.clientspecific.(current.company_str4).overview = cell(1,1);
        invoice.clientspecific.(current.company_str4).fullname = current.company_str2;   
        invoice.clientspecific.(current.company_str4).fullname2 = current.company_str3;  
        invoice.clientspecific.(current.company_str4).currency = char(UPSfile_customer(cell2mat(current.company(2,1)),cell2mat(Stringstofetch_ups(2,2))));
        invoice.clientspecific.(current.company_str4).invoicemonthly = cell2mat(UPSfile_customer(cell2mat(current.company(2,1)),cell2mat(Stringstofetch_ups(2,8))));
        invoice.clientspecific.(current.company_str4).invoicealternatively = cell2mat(UPSfile_customer(cell2mat(current.company(2,1)),cell2mat(Stringstofetch_ups(2,9))));
        invoice.clientspecific.(current.company_str4).summary = struct;
        [invoice] = addtospecificinvoice(current.company_str4,invoice,y);
    end
    % Add to the summary
    cntr = invoice.clientspecific.(current.company_str4).counter;
    %invoice.clientspecific.(current.company_str4).summary(cntr,1) = cellstr([char(invoice.clientspecific.(current.company_str4).overview(cntr,5)) ' ' char(invoice.clientspecific.(current.company_str4).overview(cntr,6))]);
    if strcmp(invoice.clientspecific.(current.company_str4).overview(cntr,6),'') || strcmp('#null',char(invoice.clientspecific.(current.company_str4).overview(cntr,6))) == 1
        currenttype = char(invoice.clientspecific.(current.company_str4).overview(cntr,5));
    else        
        currenttype = [char(invoice.clientspecific.(current.company_str4).overview(cntr,5)) ' ' char(invoice.clientspecific.(current.company_str4).overview(cntr,6))];
    end
    
    currenttype2 = regexprep(currenttype,' ','_');
    if isfield(invoice.clientspecific.(current.company_str4).summary,currenttype2) == 0
        invoice.clientspecific.(current.company_str4).summary.(currenttype2) = struct;
        invoice.clientspecific.(current.company_str4).summary.(currenttype2).name = currenttype;        
        invoice.clientspecific.(current.company_str4).summary.(currenttype2).price = cell2mat(invoice.clientspecific.(current.company_str4).overview(cntr,10));
        invoice.clientspecific.(current.company_str4).summary.(currenttype2).vat = cell2mat(invoice.clientspecific.(current.company_str4).overview(cntr,11));
        invoice.clientspecific.(current.company_str4).summary.(currenttype2).counter = 1;
    else
        invoice.clientspecific.(current.company_str4).summary.(currenttype2).counter = invoice.clientspecific.(current.company_str4).summary.(currenttype2).counter + 1;
    end
    clear current
    clear price
end
clear y

% Write to file per customer
for z = 1:invoice.nrofcustomers
    displaytext = ['Creating invoices - please wait - processing ' num2str(z) ' of ' num2str(invoice.nrofcustomers)];
    disp(displaytext);
    % Get customer field name
    currentclient = char(invoice.customerlist(z,1));   
    % Get the details
    invoicedetails = invoice.clientspecific.(currentclient);
    % Write the file
    [values] = writeinvoicetofile(invoicedetails,values);
end
clear z

% add invoice number to filtered list here.

save Temp\invoice.mat invoice

end