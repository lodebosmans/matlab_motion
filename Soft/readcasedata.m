function readcasedata()

% Written in a very bad way due to many additional requirements that were added along the way. Many lines are hardcoded.
% Be very carefull when adjusting this script. But hey, it works!

disp('Reading case data - please wait')

load Temp\values.mat values
load Temp\caseids_production.mat caseids_production
%load Temp\UPSfile_customer.mat UPSfile_customer

% load Temp\onlinereport.mat onlinereport
% load Temp\customers.mat customers
% load Temp\UPSfile_customer.mat UPSfile_customer

% Remove non-existing case IDs when running the daily snapshot
caseids_production(ismember(caseids_production,'RS19-PHI-ITS')) = [];

% Remove case IDs for which the finishing report and label is already present
if values.FinRepLabels == 1 || values.UpdateOnlineReport == 1
    for cid_idx = 1:size(caseids_production,1)
        % Get the case ID
        currentcaseid = char(caseids_production(cid_idx,1));
        currentcaseid_slim = strrep(currentcaseid,'-','');
        finrep_filename = [values.finrepfolder  currentcaseid_slim '.docx'];
        % Check if there is a finishing report for this case ID
        if isfile(finrep_filename) == 1
            % If present, remove the case ID from the list
            caseids_production(ismember(caseids_production,currentcaseid)) = {'Delete'};
        end
    end
end
caseids_production(ismember(caseids_production,'Delete')) = [];
caseids_production_tocreate = caseids_production;

nrofcases = size(caseids_production,1);
if nrofcases > 0
    load Temp\onlinereport.mat onlinereport
    load Temp\customers.mat customers
    load Temp\UPSfile_customer.mat UPSfile_customer
    
    % Get de column colnr_OR of the colums of interest
    stringstofetch = {'CaseCode','ReferenceID','HospitalName','DeliveryOfficeName','DeliveryOfficeAddress','DeliveryPostalCode','DeliveryOfficeCityName','DeliveryOfficeState','DeliveryOfficeCountry','SurgeryDate',...
        'Top_thickness','Top_size','Top_hardness','ServiceType','InsoleType','Remarks','CustomerFirstName','CustomerSurname','Heel_pad___left','Heel_pad___right',...
        'Surgeon','CaseCanceled','Top_material','BuiltActor','ShippedYear','ShippedMonth','ShippedDay','Base_type',...
        'CreatedYear','CreatedMonth','CreatedDay','CaseId','Heel_cup___left','Heel_cup___right'};
    [stringstofetch] = catchcolumnindex(stringstofetch,onlinereport,1);
end

% stringstofetch_ups = {'CustomerNumber','DeliveryOffice','CompanyNameOMS','ShipTo','InvoiceTo','Subdelivery'};
% [stringstofetch_ups] = catchcolumnindex(stringstofetch_ups,UPSfile_customer,2);

% create db_production struct
db_production = struct;
db_production.overview = cell(1,1);

% Keep the headers up to date!!! CAVE: this variable is also adjusted afterwards (e.g. in createshipmentoverview)
db_production.headers = {'caseid.full','caseid.part1','caseid.part2','caseid.part3','company','referenceID','delivery_company','delivery_street','delivery_zip_city','delivery_state',...
    'delivery_country','estimatedshippingdate','estimatedshippingdate_short','topsize','tophardness','servicetype','insoletype','remarks','customerfirstname','customersurname',...
    'heelpadleft','heelpadright','surgeon','topthickness','casecancelled','topmaterial','builtactor','shippedyear','shippedmonth','shippedday',...
    'basetype','customernumber','createdyear','createdmonth','createdday','itemnr','subdelivery','CaseId','heelcupleft','heelcupright'};


% Create for loop to iterate over all case IDs
for rownr_overview = 1:nrofcases
    % Get the caseID from the list of caseIDs
    % ---------------------------------------
    clear currentcaseid
    currentcaseid_long = char(caseids_production(rownr_overview,1));
    clc
    disp(['Reading case data (' currentcaseid_long ') - ' num2str(rownr_overview) '/' num2str(nrofcases) ' - please wait']);
    
    if strcmp('RS22-TAR-MEH',currentcaseid_long) == 1
        test = 1;
    end
    
    % For the regular case IDs
    if length(currentcaseid_long) == 12
        currentcaseid = [currentcaseid_long(1:4) currentcaseid_long(6:8) currentcaseid_long(10:12)];
        % Get the colnr_OR of the current case ID from the online report
        idx = strcmp(currentcaseid_long,onlinereport(:,cell2mat(stringstofetch(2,1))));
        rownr_OR = find(idx==1);
        
        % OnlineReport data
        % -----------------
        %
        db_production.raw.(currentcaseid).caseid.full = currentcaseid_long;
        db_production.overview(rownr_overview,1) = cellstr(db_production.raw.(currentcaseid).caseid.full);
        %
        db_production.raw.(currentcaseid).caseid.part1 = currentcaseid_long(1:4);
        db_production.overview(rownr_overview,2) = cellstr(db_production.raw.(currentcaseid).caseid.part1);
        %
        db_production.raw.(currentcaseid).caseid.part2 = currentcaseid_long(6:8);
        db_production.overview(rownr_overview,3) = cellstr(db_production.raw.(currentcaseid).caseid.part2);
        %
        db_production.raw.(currentcaseid).caseid.part3 = currentcaseid_long(10:12);
        db_production.overview(rownr_overview,4) = cellstr(db_production.raw.(currentcaseid).caseid.part3);
        %
        db_production.raw.(currentcaseid).company = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,3)))); % Get the full string of the company
        db_production.raw.(currentcaseid).company = db_production.raw.(currentcaseid).company(:,1:end-8); % Delete the word company (final 8 characters)
        db_production.overview(rownr_overview,5) = cellstr(db_production.raw.(currentcaseid).company);
        disp(['For customer: ' char(db_production.raw.(currentcaseid).company)]);
        %
        colnr_OR = cell2mat(stringstofetch(2,2));
        colnr_overview = 6;
        [db_production] = checkifnan('referenceid',currentcaseid,db_production,onlinereport,rownr_OR,colnr_OR,rownr_overview,colnr_overview);
        %
        db_production.raw.(currentcaseid).delivery_company = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,4))));
        db_production.overview(rownr_overview,7) = cellstr(db_production.raw.(currentcaseid).delivery_company);
        %
        db_production.raw.(currentcaseid).delivery_street = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,5))));
        db_production.overview(rownr_overview,8) = cellstr(db_production.raw.(currentcaseid).delivery_street);
        %
        db_production.raw.(currentcaseid).delivery_zip_city = [num2str(cell2mat(onlinereport(rownr_OR,cell2mat(stringstofetch(2,6))))) ' ' num2str(cell2mat(onlinereport(rownr_OR,cell2mat(stringstofetch(2,7)))))];
        db_production.overview(rownr_overview,9) = cellstr(db_production.raw.(currentcaseid).delivery_zip_city);
        %
        db_production.raw.(currentcaseid).delivery_state = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,8))));
        db_production.overview(rownr_overview,10) = cellstr(db_production.raw.(currentcaseid).delivery_state);
        %
        db_production.raw.(currentcaseid).delivery_country = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,9))));
        db_production.overview(rownr_overview,11) = cellstr(db_production.raw.(currentcaseid).delivery_country);
        %
        db_production.raw.(currentcaseid).estimatedshippingdate = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,10))));
        db_production.overview(rownr_overview,12) = cellstr(db_production.raw.(currentcaseid).estimatedshippingdate);
        %
        %db_production.raw.(currentcaseid).estimatedshippingdate_short = [char(db_production.raw.(currentcaseid).estimatedshippingdate(1:4)) ' / ' char(db_production.raw.(currentcaseid).estimatedshippingdate(6:7)) ' / ' char(db_production.raw.(currentcaseid).estimatedshippingdate(9:10))];
        db_production.raw.(currentcaseid).estimatedshippingdate_short = db_production.raw.(currentcaseid).estimatedshippingdate;
        db_production.overview(rownr_overview,13) = cellstr(db_production.raw.(currentcaseid).estimatedshippingdate_short);
        %
        db_production.raw.(currentcaseid).topsize = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,12))));
        db_production.overview(rownr_overview,14) = cellstr(db_production.raw.(currentcaseid).topsize);
        %
        db_production.raw.(currentcaseid).tophardness = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,13))));
        db_production.overview(rownr_overview,15) = cellstr(db_production.raw.(currentcaseid).tophardness);
        %
        db_production.raw.(currentcaseid).servicetype = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,14))));
        db_production.overview(rownr_overview,16) = cellstr(db_production.raw.(currentcaseid).servicetype);
        %
        db_production.raw.(currentcaseid).insoletype = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,15))));
        db_production.overview(rownr_overview,17) = cellstr(db_production.raw.(currentcaseid).insoletype);
        %
        db_production.raw.(currentcaseid).remarks = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,16))));
        db_production.overview(rownr_overview,18) = cellstr(db_production.raw.(currentcaseid).remarks);
        %
        temp = onlinereport(rownr_OR,cell2mat(stringstofetch(2,17)));
        if isa(temp{1},'double') == 1
            temp = num2str(cell2mat(temp));
        else
            temp = char(temp);
        end
        db_production.raw.(currentcaseid).customerfirstname = char(temp);
        db_production.overview(rownr_overview,19) = cellstr(db_production.raw.(currentcaseid).customerfirstname);
        %
        temp = onlinereport(rownr_OR,cell2mat(stringstofetch(2,18)));
        if isa(temp{1},'double') == 1
            temp = num2str(cell2mat(temp));
        elseif temp{1} == 1
            temp = 'True';
        else
            temp = char(temp);
        end
        db_production.raw.(currentcaseid).customersurname = char(temp);
        db_production.overview(rownr_overview,20) = cellstr(db_production.raw.(currentcaseid).customersurname);
        %
        [temp] = catchcolumnindex({'Heel_pad___left'},onlinereport,1);
        temp = cell2mat(temp(2,1));
        db_production.raw.(currentcaseid).heelpadleft = char(onlinereport(rownr_OR,temp));
        db_production.overview(rownr_overview,21) = cellstr(db_production.raw.(currentcaseid).heelpadleft);
        %
        [temp] = catchcolumnindex({'Heel_pad___right'},onlinereport,1);
        temp = cell2mat(temp(2,1));
        db_production.raw.(currentcaseid).heelpadright = char(onlinereport(rownr_OR,temp));
        db_production.overview(rownr_overview,22) = cellstr(db_production.raw.(currentcaseid).heelpadright);
        %
        db_production.raw.(currentcaseid).surgeon = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,21))));
        db_production.overview(rownr_overview,23) = cellstr(db_production.raw.(currentcaseid).surgeon);
        %
        db_production.raw.(currentcaseid).topthickness = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,11))));
        db_production.overview(rownr_overview,24) = cellstr(db_production.raw.(currentcaseid).topthickness);
        %
        %     heelpads = {'Plantar Fasciitis','Heel Spur','Fat Pad'};
        %     if sum(strcmp(db_production.raw.(currentcaseid).heelpadleft,heelpads)) + sum(strcmp(db_production.raw.(currentcaseid).heelpadright,heelpads)) > 0
        %         db_production.raw.(currentcaseid).topthickness = "4 mm + 2 mm";
        %         db_production.overview(rownr_overview,24) = cellstr(db_production.raw.(currentcaseid).topthickness);
        %     end
        %
        %db_production.raw.(currentcaseid).casecancelled = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,22))));
        colnr_OR = cell2mat(stringstofetch(2,22));
        colnr_overview = 25;
        [db_production] = checkifnan('casecancelled',currentcaseid,db_production,onlinereport,rownr_OR,colnr_OR,rownr_overview,colnr_overview);
        %
        colnr_OR = cell2mat(stringstofetch(2,23));
        colnr_overview = 26;
        %db_production.raw.(currentcaseid).topmaterial = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,23))));
        %db_production.overview(rownr_overview,26) = cellstr(db_production.raw.(currentcaseid).topmaterial);
        [db_production] = checkifnan('topmaterial',currentcaseid,db_production,onlinereport,rownr_OR,colnr_OR,rownr_overview,colnr_overview);
        %
        %db_production.raw.(currentcaseid).builtactor = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,24))));
        colnr_OR = cell2mat(stringstofetch(2,24));
        colnr_overview = 27;
        [db_production] = checkifnan('builtactor',currentcaseid,db_production,onlinereport,rownr_OR,colnr_OR,rownr_overview,colnr_overview);
        %
        %db_production.raw.(currentcaseid).shippedyear = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,25))));
        colnr_OR = cell2mat(stringstofetch(2,25));
        colnr_overview = 28;
        [db_production] = checkifnan('shippedyear',currentcaseid,db_production,onlinereport,rownr_OR,colnr_OR,rownr_overview,colnr_overview);
        %
        %db_production.raw.(currentcaseid).shippedmonth = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,26))));
        colnr_OR = cell2mat(stringstofetch(2,26));
        colnr_overview = 29;
        [db_production] = checkifnan('shippedmonth',currentcaseid,db_production,onlinereport,rownr_OR,colnr_OR,rownr_overview,colnr_overview);
        %
        %db_production.raw.(currentcaseid).shippedday = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,27))));
        colnr_OR = cell2mat(stringstofetch(2,27));
        colnr_overview = 30;
        [db_production] = checkifnan('shippedday',currentcaseid,db_production,onlinereport,rownr_OR,colnr_OR,rownr_overview,colnr_overview);
        %
        db_production.raw.(currentcaseid).basetype = char(onlinereport(rownr_OR,cell2mat(stringstofetch(2,28))));
        db_production.overview(rownr_overview,31) = cellstr(db_production.raw.(currentcaseid).basetype);
        %
        % CAVE: Small skip in numbering due to reorganisation script.
        %
        colnr_OR = cell2mat(stringstofetch(2,29));
        colnr_overview = 33;
        [db_production] = checkifnan('createdyear',currentcaseid,db_production,onlinereport,rownr_OR,colnr_OR,rownr_overview,colnr_overview);
        %
        colnr_OR = cell2mat(stringstofetch(2,30));
        colnr_overview = 34;
        [db_production] = checkifnan('createdmonth',currentcaseid,db_production,onlinereport,rownr_OR,colnr_OR,rownr_overview,colnr_overview);
        %
        colnr_OR = cell2mat(stringstofetch(2,31));
        colnr_overview = 35;
        [db_production] = checkifnan('createdday',currentcaseid,db_production,onlinereport,rownr_OR,colnr_OR,rownr_overview,colnr_overview);
        %
        [temp] = catchcolumnindex({'CaseId'},onlinereport,1);
        temp = cell2mat(temp(2,1));
        db_production.raw.(currentcaseid).caseidmtls = num2str(cell2mat(onlinereport(rownr_OR,temp)));
        db_production.overview(rownr_overview,38) = cellstr(db_production.raw.(currentcaseid).caseidmtls);
        %
        [temp] = catchcolumnindex({'Heel_cup___left'},onlinereport,1);
        temp = cell2mat(temp(2,1));
        db_production.raw.(currentcaseid).heelcupleft = char(onlinereport(rownr_OR,temp));
        db_production.overview(rownr_overview,39) = cellstr(db_production.raw.(currentcaseid).heelcupleft);
        %
        [temp] = catchcolumnindex({'Heel_cup___right'},onlinereport,1);
        temp = cell2mat(temp(2,1));
        db_production.raw.(currentcaseid).heelcupright = char(onlinereport(rownr_OR,temp));
        db_production.overview(rownr_overview,40) = cellstr(db_production.raw.(currentcaseid).heelcupright);
        
        
        
        % UPS file data
        % -------------
        %         % Get the customer number from the UPS file - delivery address line
        %         newcell(1,1) = cellstr(db_production.raw.(currentcaseid).delivery_company);
        %         %customerrow = catchrowindex(newcell,UPSfile_customer,cell2mat(stringstofetch_ups(2,2)));
        %         [customerrow] = catchrowindexall(newcell,UPSfile_customer,cell2mat(stringstofetch_ups(2,2)));
        %         % Get the customer number from the UPS file - company address line (multiple possibilities)
        %         newcell(1,1) = onlinereport(rownr_OR,cell2mat(stringstofetch(2,3)));
        %         [allcompanymatches] = catchrowindexall(newcell,UPSfile_customer,cell2mat(stringstofetch_ups(2,3)));
        %         % Find the match between delivery and company
        %         idx = ismember(allcompanymatches,customerrow);
        %         match = find(idx==1);
        %         matchrow = allcompanymatches(match,1);
        %         Write customernumber to db
        %         db_production.raw.(currentcaseid).customernumber = char(UPSfile_customer(matchrow,cell2mat(stringstofetch_ups(2,1))));
        %         db_production.overview(rownr_overview,32) = cellstr(db_production.raw.(currentcaseid).customernumber);
        cusdata.company = [db_production.raw.(currentcaseid).company ' Company'];
        cusdata.delivery = db_production.raw.(currentcaseid).delivery_company;
        [cnr] = getcustomernumber(cusdata,UPSfile_customer);
        cnr_ = strrep(cnr,'-','_');
        db_production.raw.(currentcaseid).customernumber = cnr;
        db_production.overview(rownr_overview,32) = cellstr(db_production.raw.(currentcaseid).customernumber);
        
        % Itemnumber
        % ----------
        headersOR = onlinereport(1,:);
        lineOR = onlinereport(rownr_OR,:);
        [itemnr] = getitemnumber(currentcaseid_long,headersOR,customers,UPSfile_customer,lineOR); % last var is optional
        
        db_production.raw.(currentcaseid).itemnr = itemnr;
        db_production.overview(rownr_overview,36) = {itemnr};
        
        % Check if it is a subdelivery or not
        %db_production.raw.(currentcaseid).subdelivery = char(UPSfile_customer(matchrow,cell2mat(stringstofetch_ups(2,6))));
        db_production.raw.(currentcaseid).subdelivery = customers.(cnr_).Subdelivery;
        db_production.overview(rownr_overview,37) = cellstr(db_production.raw.(currentcaseid).subdelivery);
        
        
        
        
        
        % For RS-numbers
    elseif length(currentcaseid_long) == 13
        % currentcaseid = RS-number in this loop
        currentcaseid = [currentcaseid_long(1:2) currentcaseid_long(4:7) currentcaseid_long(9:13)];
        db_production.raw.(currentcaseid).caseid.full = currentcaseid_long;
        db_production.overview(rownr_overview,1) = cellstr(db_production.raw.(currentcaseid).caseid.full);
        % Add dashed lines everywhere
        db_production.overview(rownr_overview,2:end) = cellstr('-');
        
        %       % First temporary solution: Add the RS Print customer-number, so that in the sort out Matlab will come and ask for who the shipment is.
        %       db_production.overview(rownr_overview,32) = cellstr('C0211');
        
        % Second final solution: read the RS-number file and fill in only the necessary information
        [brol1,brol2,cnr] = xlsread([values.rsfilesfolder currentcaseid_long '.xlsx'],1,'B3:B3');
        cnr = char(cnr);
        cnr_ = strrep(cnr,'-','_');
        
        db_production.raw.(currentcaseid).customernumber = cnr;
        db_production.overview(rownr_overview,32) = cellstr(db_production.raw.(currentcaseid).customernumber);
        
        db_production.raw.(currentcaseid).subdelivery = customers.(cnr_).Subdelivery;
        db_production.overview(rownr_overview,37) = cellstr(db_production.raw.(currentcaseid).subdelivery);
        
        db_production.raw.(currentcaseid).company = customers.(cnr_).Company;
        db_production.overview(rownr_overview,5) = cellstr(db_production.raw.(currentcaseid).company);
        
        db_production.raw.(currentcaseid).caseidmtls = '';
        db_production.overview(rownr_overview,38) = cellstr(db_production.raw.(currentcaseid).caseidmtls);
        
    end
end

% if values.FinRepLabels ==1
%     % Sort the cases based on creation date (which is parallel to the estimated shipment date), to align the cases chronologically.
%
%     col_createdyear = catchcolumnindex({'createdyear'},db_production.headers,1);
%     col_createdyear = cell2mat(col_createdyear(2,1));
%     col_createdmonth = catchcolumnindex({'createdmonth'},db_production.headers,1);
%     col_createdmonth = cell2mat(col_createdmonth(2,1));
%     col_createdday = catchcolumnindex({'createdday'},db_production.headers,1);
%     col_createdday = cell2mat(col_createdday(2,1));
%     db_production.overview = sortrows(db_production.overview,[col_createdyear,col_createdmonth,col_createdday]);
% end

if values.FinRepLabels == 1 && nrofcases > 0
    % Sort the cases based on alphabetical order
    
    col_caseid = catchcolumnindex({'caseid.full'},db_production.headers,1);
    col_caseid = cell2mat(col_caseid(2,1));
    db_production.overview = sortrows(db_production.overview,col_caseid);
    
end

clc
save Temp\db_production.mat db_production
if nrofcases > 0
    save Temp\stringstofetch.mat stringstofetch
end
save Temp\caseids_production_tocreate caseids_production_tocreate

end