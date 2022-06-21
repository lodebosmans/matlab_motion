function readUPSfile(tabnr)

load Temp\values.mat values

% Copy the UPS file with all customer information
try
    copyfile([values.upsfilepath values.upsfilename],'Temp\UPSfile.xlsx');
    copyfile([values.upsfilepath values.upsfilename],[values.pollespowerbifolder 'UpsExcel_VoorDePolle.xlsx']);
catch
    warning('Cannot open the UPS file, probably the file is opened somewhere on some computer.');
    clear all
end
%copyfile('O:\StandardOperations\UPS\RS Print\20170802_OverviewShipments_v5.xlsx','Temp\UPSfile.xlsx');

% Open UPS file with all customers
% [data.num,data.txt,UPSfile] = xlsread('Temp\UPSfile.xlsx',tabnr);

if tabnr == 1
    disp('Reading UPS file - customer details - please wait');
    [data.num,data.txt,UPSfile_customer] = xlsread('Temp\UPSfile.xlsx',tabnr);
    % Filter out the NaNs
    UPSfile_customer(cellfun(@(x) isnumeric(x) && isnan(x), UPSfile_customer)) = {''};
    save Temp\UPSfile_customer.mat UPSfile_customer;
elseif tabnr == 2
    disp('Reading UPS file - shipment details - please wait');
    [data.num,data.txt,UPSfile_shipment] = xlsread('Temp\UPSfile.xlsx',tabnr);
    % Filter out the NaNs
    UPSfile_shipment(cellfun(@(x) isnumeric(x) && isnan(x), UPSfile_shipment)) = {'-'};
    save Temp\UPSfile_shipment.mat UPSfile_shipment;
elseif tabnr == 4
    disp('Reading UPS file - product details - please wait');
    [data.num,data.txt,UPSfile_product] = xlsread('Temp\UPSfile.xlsx',tabnr);
    UPSfile_product = UPSfile_product(1:3,2:end);
%     % Filter out the NaNs
%     UPSfile_product(cellfun(@(x) isnumeric(x) && isnan(x), UPSfile_product)) = {'-'};
    save Temp\UPSfile_product.mat UPSfile_product;
elseif tabnr == 5
    disp('Reading UPS file - itemnumbers - please wait');
    [data.num,data.txt,UPSfile_itemnumbers] = xlsread('Temp\UPSfile.xlsx',tabnr);
    % Filter out the NaNs
    UPSfile_itemnumbers(cellfun(@(x) isnumeric(x) && isnan(x), UPSfile_itemnumbers)) = {'-'};
    save Temp\UPSfile_itemnumbers.mat UPSfile_itemnumbers;
% elseif tabnr == 6
%     disp('Reading UPS file - customer agreement details - please wait');
%     [data.num,data.txt,UPSfile_customagreements] = xlsread('Temp\UPSfile.xlsx',tabnr);
%     % Filter out the NaNs
%     UPSfile_customagreements(cellfun(@(x) isnumeric(x) && isnan(x), UPSfile_customagreements)) = {'-'};
%     save Temp\UPSfile_customagreements.mat UPSfile_customagreements;
end

end