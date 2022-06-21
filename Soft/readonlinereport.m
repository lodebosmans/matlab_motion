function readonlinereport()

load Temp\values.mat values

disp('Reading online report - please wait');

try
    files = dir(values.onlinereport);
    directory.onlinereport = [files(1).folder '\' files(1).name];
    copyfile(directory.onlinereport,'Temp\OnlineReport.xlsx');
catch
    warning('Fatality: cannot open the OnlineReport!');
    clear all
end
tic
[data.num,data.txt,onlinereport] = xlsread('Temp\OnlineReport.xlsx',1);
toc

% % Save a slim version of the online report
% ORslim_filename = 'OnlineReport_VoorDePolle.xlsx';
% ORslim_path = [values.pollespowerbifolder ORslim_filename];
% xlswrite(ORslim_path,onlinereport);

% Replace all with a dashed line
onlinereport(cellfun(@(x) isnumeric(x) && isnan(x), onlinereport)) = {'-'};

clear files

save Temp\onlinereport.mat onlinereport
save Temp\data.mat data
load Temp\customers.mat customers





% % Trim the file for GO4D
% % ----------------------
% disp('Trimming the online report for GO4D - please wait');
% cusname = 'go4d';
% OR_go4d = onlinereport;
% % Delete the first part of the file, to speed things up
% OR_go4d(2:16500,:) = [];
% [nrofrowsor,nrofcolsor] = size(OR_go4d);
% headers = {'CaseCode','CreationDate','CustomerSurname','CustomerFirstName','CaseCanceled','CaseStatus','Surgeon', ...
%            'HospitalName','HospitalAccountNumber','DeliveryOfficeName','DeliveryOfiiceAccountNumber','ReferenceID',...
%            'ServiceType','Shipped','InsoleType','Top_thickness','Top_size',...
%            'Top_hardness','Top_material','Heel_pad___left','Heel_pad___right','Remarks',...
%            'Base_type','CreatedYear','CreatedMonth','CreatedDay','ShippedYear',...
%            'ShippedMonth','ShippedDay','SurgeryType','Built','Shipped',...
%            'PatientBirthDate'};
% nrofheaders = size(headers,2);
% 
% % Delete the unwanted columns
% deleted = 0;
% for x = 1:nrofcolsor
%     ch = char(OR_go4d(1,x-deleted));
%     if isempty(find(contains(headers,ch))) == 1
%         % Delete column
%         OR_go4d(:,x-deleted) = [];
%         deleted = deleted + 1;
%     else
% %         % Do nothing
% %         if strcmp(ch,'CustomerSurname') == 1
% %             OR_go4d(:,x-deleted) = cellstr('CustomerSurname');
% %         end
% %         if strcmp(ch,'CustomerFirstName') == 1 
% %             OR_go4d(:,x-deleted) = cellstr('CustomerFirstName');
% %         end
%     end
% end
% 
% % Delete the unwanted rows
% col_hosname = catchcolumnindex({'HospitalName'},OR_go4d,1);
% col_hosname = cell2mat(col_hosname(2,1));
% deleted = 0;
% for x = 2:nrofrowsor
%     ch = char(OR_go4d(x-deleted,col_hosname));
%     if contains(ch,'4D') == 0 && contains(ch,'4d') == 0 && contains(ch,'4-D') == 0 && contains(ch,'4-d') == 0
%         % Delete row
%         OR_go4d(x-deleted,:) = [];
%         deleted = deleted + 1;
%     else
%         % Do nothing
%     end
% end
% 
% save Temp\OR_go4d.mat OR_go4d
% clear cusname
% 
% 
% 
% 
% 
% 
% % Trim the file for Flowbuilt
% % ---------------------------
% disp('Trimming the online report for Flowbuilt - please wait');
% cusname = 'flowbuilt';
% OR_flowbuilt = onlinereport;
% % Delete the first part of the file, to speed things up
% OR_flowbuilt(2:35500,:) = [];
% [nrofrowsor,nrofcolsor] = size(OR_flowbuilt);
% headers = {'CaseCode','CreationDate','CaseCanceled','CaseRestored','CaseStatus',...
%            'HospitalName', 'HospitalRegion', 'HospitalState', 'HospitalAddress', 'HospitalCountry', 'HospitalCityName', 'HospitalPostalCode', 'HospitalAccountNumber',...
%            'DeliveryOfficeName', 'DeliveryOfficeAddress', 'DeliveryOfficeState', 'DeliveryOfficeCountry', 'DeliveryOfficeRegion', 'DeliveryOfficeCityName', 'DeliveryPostalCode', 'DeliveryOfiiceAccountNumber',...
%            'ReferenceID',...
%            'ServiceType','InsoleType','Top_thickness','Top_size',...
%            'Top_hardness','Top_material','Heel_pad___left','Heel_pad___right','Remarks',...
%            'Base_type','SurgeryType','SurgeryDate','HospitalCountry', ...
%            'Built','Production','ReadyToShip','Shipped'};
% nrofheaders = size(headers,2);
% 
% % Delete the unwanted columns
% deleted = 0;
% for x = 1:nrofcolsor
%     ch = char(OR_flowbuilt(1,x-deleted));
%     if isempty(find(contains(headers,ch))) == 1
%         % Delete column
%         OR_flowbuilt(:,x-deleted) = [];
%         deleted = deleted + 1;
%     else
%         % Do nothing
%     end
% end
% 
% % Delete the unwanted rows
% col_hospcountry = catchcolumnindex({'HospitalCountry'},OR_flowbuilt,1);
% col_hospcountry = cell2mat(col_hospcountry(2,1));
% deleted = 0;
% for x = 2:nrofrowsor
%     ch = char(OR_flowbuilt(x-deleted,col_hospcountry));
%     if contains(ch,'United States') == 0 && contains(ch,'Canada') == 0 
%         % Delete row
%         OR_flowbuilt(x-deleted,:) = [];
%         deleted = deleted + 1;
%     else
%         % Do nothing
%     end
% end
% 
% save Temp\OR_flowbuilt.mat OR_flowbuilt
% clear cusname


% Copy the online report .mat file to the online report folder
copyfile('Temp\onlinereport.mat',values.onlinereportmat);
%copyfile('Temp\OR_go4d.mat',[values.onlinereportfolder 'OR_go4d.mat']);
%copyfile('Temp\OR_flowbuilt.mat',[values.onlinereportfolder 'OR_flowbuilt.mat']);
       
end