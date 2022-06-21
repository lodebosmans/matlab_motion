function [cnr] = getcustomernumber(cusdata,UPSfile_customer)

% load Temp\UPSfile_customer.mat UPSfile_customer

% if strcmp(cusdata.company,'Materialise HQ Company') == 1
%     cnr = 'C0202';
% elseif strcmp(cusdata.company,'SJV Medical Products Company') == 1
%     cnr = 'C0015';
% elseif strcmp(cusdata.company,'Munkebo Fodterapi Bente Frederiksen Company') == 1
%     cnr = 'C1157';
% elseif strcmp(cusdata.company,'OS Funesco GCV Company') == 1
%     cnr = 'C1031';
% else
%     if length(cusdata.delivery) > 9
%         if strcmp(cusdata.delivery(1:9),'OBSOLETE_') == 1
%             if strcmp(cusdata.delivery(10:end),'Be Balanced Be Fit') == 1
%                 cusdata.delivery = 'Be Balanced Be Fit BVBA';
%             else
%                 cusdata.delivery = cusdata.delivery(10:end);
%             end
%         end
%     end
%     [cnr] = getcustomernumber(cusdata);
%     cnr = strrep(cnr,'-','_');
% end

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% col_companynameoms = catchcolumnindex({'CompanyNameOMS'},UPSfile_customer,2);
% col_companynameoms = cell2mat(col_companynameoms(2,1));
% 
% col_deliveryoffice = catchcolumnindex({'DeliveryOffice'},UPSfile_customer,2);
% col_deliveryoffice = cell2mat(col_deliveryoffice(2,1));
% 
% companyrows = catchrowindex2({cusdata.company},UPSfile_customer,col_companynameoms);
% deliveryrows = catchrowindex2({cusdata.delivery},UPSfile_customer,col_deliveryoffice);
% 
% [~,idx]=ismember(companyrows,deliveryrows);
% finalmatch = find(idx==1);
% 
% cusnr = char(UPSfile_customer(companyrows(finalmatch,1),1)); %#ok<FNDSB>

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

if strcmp(cusdata.company,'Materialise HQ Company') == 1
    cnr = 'C0202';
elseif strcmp(cusdata.company,'SJV Medical Products Company') == 1
    cnr = 'C0015';
elseif strcmp(cusdata.company,'Munkebo Fodterapi Bente Frederiksen Company') == 1
    cnr = 'C1157';
elseif strcmp(cusdata.company,'OS Funesco GCV Company') == 1
    cnr = 'C1031';
elseif strcmp(cusdata.company,'Atelier D Orthopedie Company') == 1
    cnr = 'C1037';
elseif strcmp(cusdata.company,'RS Print Company') == 1
    cnr = 'C0210';
elseif strcmp(cusdata.company,'Podologie Rios Company') == 1 && strcmp(cusdata.delivery,'Otkas BVBA') == 1
    cnr = 'C1219';
elseif strcmp(cusdata.company,'company1') == 1 && strcmp(cusdata.delivery,'ofice1') == 1
    cnr = 'C0202';
elseif strcmp(cusdata.company,'Company') == 1 && strcmp(cusdata.delivery,'Ofice') == 1
    cnr = 'C0202';
elseif strcmp(cusdata.company,'Military Hopsital Queen Astrid Company') == 1 && strcmp(cusdata.delivery,'Military Hopsital Queen Astrid Company') == 1
    cnr = 'C1122';
else
    if length(cusdata.delivery) > 9
        if strcmp(cusdata.delivery(1:9),'OBSOLETE_') == 1
            if strcmp(cusdata.delivery(10:end),'Be Balanced Be Fit') == 1
                cusdata.delivery = 'Be Balanced Be Fit BVBA';
            elseif strcmp(cusdata.delivery(10:end),'Atelier D Orthopedie') == 1
                cusdata.delivery = 'Atelier D Orthopedie';
            else
                cusdata.delivery = cusdata.delivery(10:end);
            end
        end
    end
    
    col_companynameoms = catchcolumnindex({'CompanyNameOMS'},UPSfile_customer,2);
    col_companynameoms = cell2mat(col_companynameoms(2,1));
    
    col_deliveryoffice = catchcolumnindex({'DeliveryOffice'},UPSfile_customer,2);
    col_deliveryoffice = cell2mat(col_deliveryoffice(2,1));
    
    companyrows = catchrowindex2({cusdata.company},UPSfile_customer,col_companynameoms);
    deliveryrows = catchrowindex2({cusdata.delivery},UPSfile_customer,col_deliveryoffice);
    
    [~,idx]=ismember(companyrows,deliveryrows);
    finalmatch = find(idx==1);
    
    cnr = char(UPSfile_customer(companyrows(finalmatch,1),1)); %#ok<FNDSB>
    cnr = strrep(cnr,'-','_');
end

end