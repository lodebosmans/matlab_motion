function readUPSfileUS(values,tabnr)

try
    files = dir('O:\StandardOperations\UPS\RS Print\US\*OverviewShipments_US*.xlsx');
    directory.UPSfileUS = [files(1).folder '\' files(1).name];
    copyfile(directory.UPSfileUS,'Temp\UPSfileUS.xlsx');
catch
    warning('Fatality: cannot open the UPS file US!');
    clear all
end

[data.num,data.txt,UPSfileUS] = xlsread('Temp\UPSfileUS.xlsx',tabnr);

UPSfileUS(cellfun(@(x) isnumeric(x) && isnan(x), UPSfileUS)) = {'-'};

clear files

save Temp\UPSfileUS.mat UPSfileUS

end