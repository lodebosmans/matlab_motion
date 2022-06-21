function readonlinereport3()

% Reads the online report and caculates the price of the case. Does not update the itemnumber in the proces
disp('Checking order statusses - updating if necessary - please wait');

load Temp\values.mat values
load Temp\customers.mat customers
load Temp\onlinereport.mat onlinereport
load([values.salesordercheck_folder 'salesordercheck_v3.mat'],'salesordercheck_v3');
load Temp\UPSfile_customer.mat UPSfile_customer

% Get the headers out of the OR
headersOR = onlinereport(1,:);

% col_year = catchcolumnindex({'CreatedYear'},headersOR,1);
% col_year = cell2mat(col_year(2,1));
col_CaseId = catchcolumnindex({'CaseId'},headersOR,1);
col_CaseId = cell2mat(col_CaseId(2,1));
col_CaseCode = catchcolumnindex({'CaseCode'},headersOR,1);
col_CaseCode = cell2mat(col_CaseCode(2,1));
col_CreatedYear = catchcolumnindex({'CreatedYear'},headersOR,1);
col_CreatedYear = cell2mat(col_CreatedYear(2,1));
col_CreatedMonth = catchcolumnindex({'CreatedMonth'},headersOR,1);
col_CreatedMonth = cell2mat(col_CreatedMonth(2,1));
col_CreatedDay = catchcolumnindex({'CreatedDay'},headersOR,1);
col_CreatedDay = cell2mat(col_CreatedDay(2,1));
col_HospitalName = catchcolumnindex({'HospitalName'},headersOR,1);
col_HospitalName = cell2mat(col_HospitalName(2,1));
% col_DeliveryOfficeName = catchcolumnindex({'DeliveryOfficeName'},headersOR,1);
% col_DeliveryOfficeName = cell2mat(col_DeliveryOfficeName(2,1));
col_InsoleType = catchcolumnindex({'InsoleType'},headersOR,1);
col_InsoleType = cell2mat(col_InsoleType(2,1));
col_Base_type = catchcolumnindex({'Base_type'},headersOR,1);
col_Base_type = cell2mat(col_Base_type(2,1));
col_CaseCanceled = catchcolumnindex({'CaseCanceled'},headersOR,1);
col_CaseCanceled = cell2mat(col_CaseCanceled(2,1));
col_DeliveryOfficeName = catchcolumnindex({'DeliveryOfficeName'},headersOR,1);
col_DeliveryOfficeName = cell2mat(col_DeliveryOfficeName(2,1));
col_ServiceType = catchcolumnindex({'ServiceType'},headersOR,1);
col_ServiceType = cell2mat(col_ServiceType(2,1));
col_Top_material = catchcolumnindex({'Top_material'},headersOR,1);
col_Top_material = cell2mat(col_Top_material(2,1));
col_HospitalRegion = catchcolumnindex({'HospitalRegion'},headersOR,1);
col_HospitalRegion = cell2mat(col_HospitalRegion(2,1));
col_CreationDate = catchcolumnindex({'CreationDate'},headersOR,1);
col_CreationDate = cell2mat(col_CreationDate(2,1));
col_CaseRestored = catchcolumnindex({'CaseRestored'},headersOR,1);
col_CaseRestored = cell2mat(col_CaseRestored(2,1));
col_CaseStatus = catchcolumnindex({'CaseStatus'},headersOR,1);
col_CaseStatus = cell2mat(col_CaseStatus(2,1));
col_HospitalAccountNumber = catchcolumnindex({'HospitalAccountNumber'},headersOR,1);
col_HospitalAccountNumber = cell2mat(col_HospitalAccountNumber(2,1));
col_DeliveryOfiiceAccountNumber = catchcolumnindex({'DeliveryOfiiceAccountNumber'},headersOR,1);
col_DeliveryOfiiceAccountNumber = cell2mat(col_DeliveryOfiiceAccountNumber(2,1));
col_DeliveryOfficeRegion = catchcolumnindex({'DeliveryOfficeRegion'},headersOR,1);
col_DeliveryOfficeRegion = cell2mat(col_DeliveryOfficeRegion(2,1));
col_ReferenceID = catchcolumnindex({'ReferenceID'},headersOR,1);
col_ReferenceID = cell2mat(col_ReferenceID(2,1));
col_Top_thickness = catchcolumnindex({'Top_thickness'},headersOR,1);
col_Top_thickness = cell2mat(col_Top_thickness(2,1));
col_Top_size = catchcolumnindex({'Top_size'},headersOR,1);
col_Top_size = cell2mat(col_Top_size(2,1));
col_Top_hardness = catchcolumnindex({'Top_hardness'},headersOR,1);
col_Top_hardness = cell2mat(col_Top_hardness(2,1));
col_Heel_pad___left = catchcolumnindex({'Heel_pad___left'},headersOR,1);
col_Heel_pad___left = cell2mat(col_Heel_pad___left(2,1));
col_Heel_pad___right = catchcolumnindex({'Heel_pad___right'},headersOR,1);
col_Heel_pad___right = cell2mat(col_Heel_pad___right(2,1));
col_Remarks = catchcolumnindex({'Remarks'},headersOR,1);
col_Remarks = cell2mat(col_Remarks(2,1));
col_SurgeryType = catchcolumnindex({'SurgeryType'},headersOR,1);
col_SurgeryType = cell2mat(col_SurgeryType(2,1));
col_SurgeryDate = catchcolumnindex({'SurgeryDate'},headersOR,1);
col_SurgeryDate = cell2mat(col_SurgeryDate(2,1));
col_Built = catchcolumnindex({'Built'},headersOR,1);
col_Built = cell2mat(col_Built(2,1));
col_Production = catchcolumnindex({'Production'},headersOR,1);
col_Production = cell2mat(col_Production(2,1));
col_ReadyToShip = catchcolumnindex({'ReadyToShip'},headersOR,1);
col_ReadyToShip = cell2mat(col_ReadyToShip(2,1));
col_Shipped = catchcolumnindex({'Shipped'},headersOR,1);
col_Shipped = cell2mat(col_Shipped(2,1));
col_ShippedYear = catchcolumnindex({'ShippedYear'},headersOR,1);
col_ShippedYear = cell2mat(col_ShippedYear(2,1));
col_ShippedMonth = catchcolumnindex({'ShippedMonth'},headersOR,1);
col_ShippedMonth = cell2mat(col_ShippedMonth(2,1));
col_ShippedDay = catchcolumnindex({'ShippedDay'},headersOR,1);
col_ShippedDay = cell2mat(col_ShippedDay(2,1));
col_Surgeon = catchcolumnindex({'Surgeon'},headersOR,1);
col_Surgeon = cell2mat(col_Surgeon(2,1));
col_CustomerSurname = catchcolumnindex({'CustomerSurname'},headersOR,1);
col_CustomerSurname = cell2mat(col_CustomerSurname(2,1));
col_CustomerFirstName = catchcolumnindex({'CustomerFirstName'},headersOR,1);
col_CustomerFirstName = cell2mat(col_CustomerFirstName(2,1));
col_PatientBirthDate = catchcolumnindex({'PatientBirthDate'},headersOR,1);
col_PatientBirthDate = cell2mat(col_PatientBirthDate(2,1));
col_HeelCupLeft = catchcolumnindex({'Heel_cup___left'},headersOR,1);
col_HeelCupLeft = cell2mat(col_HeelCupLeft(2,1));
col_HeelCupRight = catchcolumnindex({'Heel_cup___right'},headersOR,1);
col_HeelCupRight = cell2mat(col_HeelCupRight(2,1));




% Sort out all the orders and add the price
% -----------------------------------------

% Get the last line that was generated in the SOC
SOC_last_ID = size(salesordercheck_v3.data,1);
% Get the associated caseID
SOC_last_caseID = char(salesordercheck_v3.data(SOC_last_ID,1));
% Add 1, as this is the next ID that needs to be processed in the SOC.
SOC_startpoint_ID = SOC_last_ID + 1;

% Get the last ID in the OR
OR_endpoint_ID = cell2mat(onlinereport(end,col_CaseId));
OR_endpoint_row = size(onlinereport,1);
OR_endpoint_caseID = char(onlinereport(OR_endpoint_row,col_CaseCode));
% orrows = size(onlinereport,1);

if OR_endpoint_ID >= SOC_startpoint_ID
    % Calculate where to start reading in the OR. Find the last case ID that was processed in the SOC
    OR_last_caseID_from_SOC_row = catchrowindex({SOC_last_caseID},onlinereport,col_CaseCode);
    OR_last_caseID_from_SOC_row = cell2mat(OR_last_caseID_from_SOC_row(2,1));
    % Add one, as we want to start reading from the next one onwards
    OR_startpoint_row = OR_last_caseID_from_SOC_row + 1;
    OR_startpoint_ID = cell2mat(onlinereport(OR_startpoint_row,1));
    OR_startpoint_caseID = char(onlinereport(OR_startpoint_row,2));
    
    % Get the amount of cases that needs a finishing report
    caseids_production = onlinereport(OR_startpoint_row:OR_endpoint_row,col_CaseCode);
    save Temp\caseids_production.mat caseids_production
    % Do the loop for every case.
    for cc = OR_startpoint_row:OR_endpoint_row
        if cc == 106084
            test = 1;
        end
        lineOR = onlinereport(cc,:);
        % Get the case ID
        ccid = char(lineOR(1,col_CaseCode));
        cmtlsid = cell2mat(lineOR(1,col_CaseId));
        disp(' ');
        disp(['Case ' ccid ' - ' num2str(cmtlsid) '/' num2str(OR_endpoint_ID)]);
        salesordercheck_v3.data(cmtlsid,1) = {ccid};
        
        cy = cell2mat(lineOR(1,col_CreatedYear));
        cm = cell2mat(lineOR(1,col_CreatedMonth));
        cd = cell2mat(lineOR(1,col_CreatedDay));
        if strcmp(ccid,'RS18-NEK-KUN') == 1
            cy = 2018;
            cm = 8;
            cd = 14;
        end
        [itemnr,cusnr] = getitemnumber(ccid,headersOR,customers,UPSfile_customer,lineOR);
        disp(['Customer number = ' cusnr]);
        if cc > 13416
            [agrnr_cus] = getcurrentpriceagreement(customers,cusnr,cy,cm,cd);
            agr_cus = customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr_cus)]);
        end
        
        if strcmp(lineOR(1,col_CaseCanceled),'-') == 0
            salesordercheck_v3.data(cmtlsid,2:17) = {'-'};
            salesordercheck_v3.data(cmtlsid,5) = {'0: Cancelled'};
        elseif strcmp(lineOR(1,col_CaseCanceled),'-') == 1 && cc < 13417
            salesordercheck_v3.data(cmtlsid,2:17) = {'-'};
            salesordercheck_v3.data(cmtlsid,5) = {'Pre-Navision Era'};
        else
            if cc == 49327
                test = 1;
            end
            
            %         if strcmp(char(lineOR(1,col_HospitalName)),'Ortho Georgia Company') == 1
            %             test = 1;
            %         end
            compname = char(lineOR(1,col_HospitalName));
            delname = char(lineOR(1,col_DeliveryOfficeName));
            disp(['For customer: ' compname ' - ' delname ]);
            % Get the creation date
            %cy = cell2mat(lineOR(1,col_CreatedYear));
            %cm = cell2mat(lineOR(1,col_CreatedMonth));
            %cd = cell2mat(lineOR(1,col_CreatedDay));
            %[itemnr,cusnr] = getitemnumber(ccid,headersOR,customers,lineOR);
            
            flow = 0;
            nrs.mainnr = '-';
            nrs.agentnr = '-';
            nrs.agentmainnr = '-';
            % Get the correct price agreement based on the creation date
            %[agrnr_cus] = getcurrentpriceagreement(customers,cusnr,cy,cm,cd);
            %agr_cus = customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr_cus)]);
            % Identify the relevant agreement
            if strcmp(customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr_cus)]).PricingMethod,'Own') == 1
                % If own princing method, use that on.
                flow = 1;
                agr_price = customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr_cus)]);
                %nrs.agentnr = customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr_cus)]).Agent;
            elseif strcmp(customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr_cus)]).PricingMethod,'Main') == 1
                % If main account pricing method, get the info from the main account.
                nrs.mainnr = customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr_cus)]).Main;
                nrs.mainnr = strrep(nrs.mainnr,'-','_');
                [agrnr_main] = getcurrentpriceagreement(customers,nrs.mainnr,cy,cm,cd);
                if strcmp(customers.(nrs.mainnr).PriceAgreements.(['Agreement' num2str(agrnr_main)]).PricingMethod,'Own') == 1
                    % If own princing method, use that on.
                    flow = 2;
                    agr_price = customers.(nrs.mainnr).PriceAgreements.(['Agreement' num2str(agrnr_main)]);
                    %nrs.agentnr = customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr_cus)]).Agent;
                elseif strcmp(customers.(nrs.mainnr).PriceAgreements.(['Agreement' num2str(agrnr_main)]).PricingMethod,'Agent') == 1
                    % If agent account pricing method, get the info from the agent account.
                    nrs.agentnr = customers.(nrs.mainnr).PriceAgreements.(['Agreement' num2str(agrnr_main)]).Agent;
                    nrs.agentnr = strrep(nrs.agentnr,'-','_');
                    [agrnr_agent] = getcurrentpriceagreement(customers,nrs.agentnr,cy,cm,cd);
                    if strcmp(customers.(nrs.agentnr).PriceAgreements.(['Agreement' num2str(agrnr_agent)]).PricingMethod,'Own') == 1
                        % If own princing method, use that on.
                        flow = 3;
                        agr_price = customers.(nrs.agentnr).PriceAgreements.(['Agreement' num2str(agrnr_agent)]);
                    elseif strcmp(customers.(nrs.agentnr).PriceAgreements.(['Agreement' num2str(agrnr_agent)]).PricingMethod,'Main') == 1
                        % If main agent account pricing method, get the info from the main account.
                        flow = 4;
                        nrs.agentmainnr = customers.(nrs.agentnr).PriceAgreements.(['Agreement' num2str(agrnr_agent)]).Main;
                        nrs.agentmainnr = strrep(nrs.agentmainnr,'-','_');
                        [agrnr_agentmain] = getcurrentpriceagreement(customers,nrs.agentmainnr,cy,cm,cd);
                        agr_price = customers.(nrs.agentmainnr).PriceAgreements.(['Agreement' num2str(agrnr_agentmain)]);
                    else
                        disp('Fatal error!');
                        clear all
                    end
                end
            elseif strcmp(customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr_cus)]).PricingMethod,'Agent') == 1
                % If agent account pricing method, get the info from the agent account.
                nrs.agentnr = customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr_cus)]).Agent;
                nrs.agentnr = strrep(nrs.agentnr,'-','_');
                [agrnr_agent] = getcurrentpriceagreement(customers,nrs.agentnr,cy,cm,cd);
                if strcmp(customers.(nrs.agentnr).PriceAgreements.(['Agreement' num2str(agrnr_agent)]).PricingMethod,'Own') == 1
                    % If own princing method, use that on.
                    flow = 5;
                    agr_price = customers.(nrs.agentnr).PriceAgreements.(['Agreement' num2str(agrnr_agent)]);
                elseif strcmp(customers.(nrs.agentnr).PriceAgreements.(['Agreement' num2str(agrnr_agent)]).PricingMethod,'Main') == 1
                    % If main agent account pricing method, get the info from the main account.
                    flow = 6;
                    nrs.agentmainnr = customers.(nrs.agentnr).PriceAgreements.(['Agreement' num2str(agrnr_agent)]).Main;
                    nrs.agentmainnr = strrep(nrs.agentmainnr,'-','_');
                    [agrnr_agentmain] = getcurrentpriceagreement(customers,nrs.agentmainnr,cy,cm,cd);
                    agr_price = customers.(nrs.agentmainnr).PriceAgreements.(['Agreement' num2str(agrnr_agentmain)]);
                else
                    disp('Fatal error!');
                    clear all
                end
            else
                disp('Fatal error!');
                clear all
            end
            
            %             % Check if this customer already has a field in the salesordercheck_v3.orders
            %             [salesordercheck_v3] = checkfororderstruct(salesordercheck_v3,cusnr,cy);
            %             % Do the same for the main/agent if necessary
            %             if ((flow == 1 || flow == 2) &&  strcmp(nrs.agentnr,'-') == 0) || flow == 3 || flow == 5
            %                 [salesordercheck_v3] = checkfororderstruct(salesordercheck_v3,nrs.agentnr,cy);
            %             elseif flow == 4 || flow == 6
            %                 [salesordercheck_v3] = checkfororderstruct(salesordercheck_v3,nrs.agentmainnr,cy);
            %             end
            
            % Check if this customer already has a field in the salesordercheck_v3.orders
            [salesordercheck_v3] = checkfororderstruct(salesordercheck_v3,cusnr,cy);
            % Do the same for the main/agent if necessary
            if strcmp(nrs.mainnr,'-') == 0
                [salesordercheck_v3] = checkfororderstruct(salesordercheck_v3,nrs.mainnr,cy);
            end
            if strcmp(nrs.agentnr,'-') == 0
                [salesordercheck_v3] = checkfororderstruct(salesordercheck_v3,nrs.agentnr,cy);
            end
            if strcmp(nrs.agentmainnr,'-') == 0
                [salesordercheck_v3] = checkfororderstruct(salesordercheck_v3,nrs.agentmainnr,cy);
            end
            
            % Define the insole type
            typenr = str2double(itemnr(3));
            if typenr == 0
                type2 = 'RegSemi';
            elseif typenr == 1
                type2 = 'RegFull';
            elseif typenr == 2
                type2 = 'SlimSemi';
            elseif typenr == 3
                type2 = 'SlimFull';
            else
                disp('Insole type not recognized');
                clear all
            end
            
            % Get the price strategy and discount cascades.
            pricingmethod.leasing = 0;
            pricingmethod.starterpack = 0;
            pricingmethod.license = 0;
            pricingmethod.standardprice = 0;
            pricingmethod.customprice = 0;
            pricingmethod.toplayerprice = 0;
            pricingmethod.splittypecounter = 0;
            % Firstly, look at the special formats
            if strcmp(agr_price.LeasingInsoleAmount,'-') == 0
                pricingmethod.leasing = 1;
            elseif strcmp(agr_price.StarterPackSize,'-') == 0
                pricingmethod.starterpack = 1;
            elseif strcmp(agr_price.LicMin,'-') == 0
                pricingmethod.license = 1;
            end
            % Secondly, look at the pricing format
            if strcmp(agr_price.StandardPrice,'Yes') == 1
                pricingmethod.standardprice = 1;
            elseif strcmp(agr_price.RegSemiCus,'-') == 0
                pricingmethod.customprice = 1;
            elseif strcmp(agr_price.TL_RegSemi_None,'-') == 0
                pricingmethod.toplayerprice = 1;
                material = char(lineOR(1,col_Top_material));
                assembly = char(lineOR(1,col_ServiceType));
                basetype = char(lineOR(1,col_Base_type));
                [type2] = getinsoletype2(itemnr,material,assembly,basetype);
            end
            if strcmp('YES',upper(agr_price.SplitTypeCounter)) == 1
                pricingmethod.splittypecounter = 1;
            end
            
            % Determine the correct price based on the relevant flow
            if pricingmethod.leasing == 1
                [endyear,endmonth,endday] = getenddate2(agr_price.LeasingStartYear,agr_price.LeasingStartMonth,agr_price.LeasingStartDay,agr_price.LeasingDuration);
                % Check if the agreement is still active
                [active] = checkifactive(agr_price.LeasingStartYear,agr_price.LeasingStartMonth,agr_price.LeasingStartDay,cy,cm,cd,endyear,endmonth,endday);
                if active == 1
                    % Check if the amount has been reached?
                    alreadyordered = size(salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).(['m' num2str(cm)]),1);
                    minimumcontract = agr_price.LeasingInsoleAmount;
                    if alreadyordered >= minimumcontract
                        % If yes, give regular pricing.
                        % -------------------------------------------------------------------------------
                        [currentprice,discount,comment] = getregularprice(salesordercheck_v3,agr_cus,agr_price,pricingmethod,flow,type2,cy,cm);
                    else
                        % If not, give 100% discount.
                        currentprice = agr_price.LeasingMinimalPrice/agr_price.LeasingInsoleAmount;
                        discount = 100;
                        comment = ['Leasing order ' num2str(alreadyordered+1) '/' num2str(minimumcontract) ' in month ' num2str(cm) '/' num2str(cy) ];
                    end
                else
                    % If no, go to regular pricing.
                    % -------------------------------------------------------------------------------
                    [currentprice,discount,comment] = getregularprice(salesordercheck_v3,agr_cus,agr_price,pricingmethod,flow,type2,cy,cm);
                end
            elseif pricingmethod.starterpack == 1
                [endyear,endmonth,endday]  = getenddate2(agr_price.StarterPackYear,agr_price.StarterPackMonth,agr_price.StarterPackDay,agr_price.StarterpackDuration);
                % Check if the agreement is still active
                [active] = checkifactive(agr_price.StarterPackYear,agr_price.StarterPackMonth,agr_price.StarterPackDay,cy,cm,cd,endyear,endmonth,endday);
                if active == 1
                    % Check if the amount has been reached?
                    [currentprice,discount,comment] = getregularprice(salesordercheck_v3,agr_cus,agr_price,pricingmethod,flow,type2,cy,cm);
                    alreadyordered = salesordercheck_v3.orders.(cusnr).totalordersever;
                    minimumcontract = agr_price.StarterPackSize;
                    if alreadyordered >= minimumcontract
                        % If yes, give regular pricing.
                        % Do nothing, price is correct.
                    else
                        % If not, give 100% discount.
                        discount = 100;
                        comment = ['Starterpack ' num2str(cm) '/' num2str(cy) ' - order ' num2str(alreadyordered+1) '/' num2str(minimumcontract)];
                    end
                else
                    % If not, go to regular pricing.
                    % -------------------------------------------------------------------------------
                    [currentprice,discount,comment] = getregularprice(salesordercheck_v3,agr_cus,agr_price,pricingmethod,flow,type2,cy,cm);
                end
            elseif pricingmethod.license == 1
                % regular pricing
                [currentprice,discount,comment] = getregularprice(salesordercheck_v3,agr_cus,agr_price,pricingmethod,flow,type2,cy,cm);
            else
                % regular pricing
                [currentprice,discount,comment] = getregularprice(salesordercheck_v3,agr_cus,agr_price,pricingmethod,flow,type2,cy,cm);
            end
            
            % Write some first data in the salesordercheck
            cusnr = strrep(cusnr,'_','-');
            salesordercheck_v3.data(cmtlsid,2) = {cusnr};
            salesordercheck_v3.data(cmtlsid,3) = {compname};
            salesordercheck_v3.data(cmtlsid,4) = {delname};
            cusnr = strrep(cusnr,'-','_');
            salesordercheck_v3.data(cmtlsid,6) = {itemnr};
            itemnr2 = char(itemnr);
            if strcmp(itemnr2(2),'X') == 1
                salesordercheck_v3.data(cmtlsid,5) = {'1: Temporary itemnumber'};
            else
                salesordercheck_v3.data(cmtlsid,5) = {'2: Final itemnumber - ready for Navision'};
            end
            salesordercheck_v3.data(cmtlsid,7:end) = {'-'};
            
            % Add the order to the list of the month and the year + update the counters
            newrowm = size(salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).(['m' num2str(cm)]),1) + 1;
            eval(['salesordercheck_v3.orders.' cusnr '.y' num2str(cy) '.m' num2str(cm) '(' num2str(newrowm) ',1) = {cmtlsid};']);
            eval(['salesordercheck_v3.orders.' cusnr '.y' num2str(cy) '.m' num2str(cm) '(' num2str(newrowm) ',2) = {[''' num2str(cd) '/' num2str(cm) '/' num2str(cy) ''']};']);
            eval(['salesordercheck_v3.orders.' cusnr '.y' num2str(cy) '.m' num2str(cm) '(' num2str(newrowm) ',3) = {ccid};']);
            eval(['salesordercheck_v3.orders.' cusnr '.y' num2str(cy) '.m' num2str(cm) '(' num2str(newrowm) ',4) = {type2};']);
            eval(['salesordercheck_v3.orders.' cusnr '.y' num2str(cy) '.m' num2str(cm) '(' num2str(newrowm) ',5) = {currentprice};']);
            eval(['salesordercheck_v3.orders.' cusnr '.y' num2str(cy) '.m' num2str(cm) '(' num2str(newrowm) ',6) = {discount};']);
            eval(['salesordercheck_v3.orders.' cusnr '.y' num2str(cy) '.m' num2str(cm) '(' num2str(newrowm) ',7) = {comment};']);
            % Do the same for the year counter
            newrowy = size(salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).overview,1) + 1;
            eval(['salesordercheck_v3.orders.' cusnr '.y' num2str(cy) '.overview(' num2str(newrowy) ',1:7) = salesordercheck_v3.orders.' cusnr '.y' num2str(cy) '.m' num2str(cm) '(' num2str(newrowm) ',1:7);']);
            % Update the type counter
            eval(['salesordercheck_v3.orders.' cusnr '.y' num2str(cy) '.Type_' type2 ' = salesordercheck_v3.orders.' cusnr '.y' num2str(cy) '.Type_' type2 ' + 1;' ])
            % And finally the total orders ever
            salesordercheck_v3.orders.(cusnr).totalordersever = salesordercheck_v3.orders.(cusnr).totalordersever + 1;
            
            % Add some date to the general overview
            % Creation data
            salesordercheck_v3.data(cmtlsid,7) = {[num2str(cd) '/' num2str(cm) '/' num2str(cy)]};
            % Price
            salesordercheck_v3.data(cmtlsid,8) = {currentprice};
            % Price comment
            salesordercheck_v3.data(cmtlsid,9) = {comment};
            
            %             if strcmp(cusnr,'C1093') == 1
            %                 test = 1;
            %             end
            
            
            % Add to agent accounts op basis van flow nummer?
            if ((flow == 1 || flow == 2) &&  strcmp(nrs.agentnr,'-') == 0) || flow == 3 || flow == 5
                newrowm_agent = size(salesordercheck_v3.orders.(nrs.agentnr).(['y' num2str(cy)]).(['Peasant_m' num2str(cm)]),1) + 1;
                salesordercheck_v3.orders.(nrs.agentnr).(['y' num2str(cy)]).(['Peasant_m' num2str(cm)])(newrowm_agent,1) = {ccid};
                newrowy_agent = size(salesordercheck_v3.orders.(nrs.agentnr).(['y' num2str(cy)]).Peasant_overview,1) + 1;
                salesordercheck_v3.orders.(nrs.agentnr).(['y' num2str(cy)]).Peasant_overview(newrowy_agent,1) = {ccid};
                salesordercheck_v3.orders.(nrs.agentnr).(['y' num2str(cy)]).(['Peasant_Type_' type2]) = salesordercheck_v3.orders.(nrs.agentnr).(['y' num2str(cy)]).(['Peasant_Type_' type2]) + 1;
                salesordercheck_v3.orders.(nrs.agentnr).totalordersever = salesordercheck_v3.orders.(nrs.agentnr).totalordersever + 1;
            elseif flow == 4 || flow == 6
                newrowm_agent = size(salesordercheck_v3.orders.(nrs.agentmainnr).(['y' num2str(cy)]).(['Peasant_m' num2str(cm)]),1) + 1;
                salesordercheck_v3.orders.(nrs.agentmainnr).(['y' num2str(cy)]).(['Peasant_m' num2str(cm)])(newrowm_agent,1) = {ccid};
                newrowy_agent = size(salesordercheck_v3.orders.(nrs.agentmainnr).(['y' num2str(cy)]).Peasant_overview,1) + 1;
                salesordercheck_v3.orders.(nrs.agentmainnr).(['y' num2str(cy)]).Peasant_overview(newrowy_agent,1) = {ccid};
                salesordercheck_v3.orders.(nrs.agentmainnr).(['y' num2str(cy)]).(['Peasant_Type_' type2]) = salesordercheck_v3.orders.(nrs.agentmainnr).(['y' num2str(cy)]).(['Peasant_Type_' type2]) + 1;
                salesordercheck_v3.orders.(nrs.agentmainnr).totalordersever = salesordercheck_v3.orders.(nrs.agentmainnr).totalordersever + 1;
            end
            
            
            
            
        end
        
        % If a North American order, add to datafeed
        % Get the surname and check the content
        temp_surname = lineOR(1,col_CustomerSurname);
        if isa(temp_surname{1},'double') == 1
            temp_surname = num2str(cell2mat(temp_surname));
        else
            temp_surname = char(temp_surname);
        end
        if (strcmp(char(lineOR(1,col_HospitalRegion)),'USA') == 1 || strcmp(char(lineOR(1,col_HospitalRegion)),'CAN') == 1) ...
                && strcmp(temp_surname,'CleanData') == 0
            
            if strcmp(cusnr,'CXXXX') == 0
                disp('Adding data to datafeed');
                
                nextentry = size(salesordercheck_v3.datafeed,1) + 1;
                
                if nextentry == 8000
                    test = 1;
                end
                
                % Get the ship to code
                if cc > 13416
                    if strcmp(agr_cus.ShipTo,'-') == 1 || strcmp(agr_cus.ShipTo,'D1160') == 1 || strcmp(agr_cus.ShipTo,'D0032') == 1 || strcmp(agr_cus.ShipTo,'D0001') == 1
                        cshiptonr = cusnr;
                    else
                        cshiptonr = strrep(agr_cus.ShipTo,'D','C');
                    end
                else
                    cshiptonr = cusnr;
                end
                cshiptonr = strrep(cshiptonr,'-','_');
                
                %             datafeed_headers = {'CaseCode','CreationDate','CaseCanceled','CaseRestored','CaseStatus',...
                %                 'HospitalAccountNumber','HospitalName','HospitalAddress','HospitalPostalCode','HospitalCityName','HospitalState','HospitalCountry','HospitalRegion',...
                %                 'DeliveryOfiiceAccountNumber','DeliveryOfficeName','DeliveryOfficeAddress','DeliveryPostalCode','DeliveryOfficeCityName','DeliveryOfficeState','DeliveryOfficeCountry','DeliveryOfficeRegion', ...
                %                 'ReferenceID','ServiceType','InsoleType','Top_thickness','Top_size',...
                %                 'Top_hardness','Top_material','Heel_pad___left','Heel_pad___right','Remarks',...
                %                 'Base_type','SurgeryType','SurgeryDate',...
                %                 'Built','Production','ReadyToShip','Shipped',...
                %                 'CreatedYear','CreatedMonth','CreatedDay','ShippedYear','ShippedMonth','ShippedDay',...
                %                 'Surgeon','PatientBirthDate','CustomerFirstName','CustomerSurname','Initials','Itemnumber','MeasurementUnit'};
                
                tobeoverwritten = {'HospitalAddress','HospitalPostalCode','HospitalCityName','HospitalState','HospitalCountry','HospitalRegion'...
                    'DeliveryOfficeAddress','DeliveryPostalCode','DeliveryOfficeCityName','DeliveryOfficeState','DeliveryOfficeCountry','DeliveryOfficeRegion','AlternativeCustomerNumber'};
                
                datafeed_headers = salesordercheck_v3.datafeed(1,:);
                for ch = 1:size(datafeed_headers,2)
                    % Get the header
                    cheader = char(datafeed_headers(1,ch));
                    
                    % Read the data present in the field
                    %                     if strcmp(cheader,'Initials') == 0 ...
                    %                             && strcmp(cheader,'Itemnumber') == 0 ...
                    %                             && strcmp(cheader,'MeasurementUnit') == 0 ...
                    %                             && sum(contains(tobeoverwritten,cheader)) == 0
                    if sum(contains({'MeasurementUnit','Itemnumber','Initials','ContactTelephone','ContactName','ContactEmail'},cheader)) == 0 ...
                            && sum(contains(tobeoverwritten,cheader)) == 0
                        
                        value = lineOR(1,eval(['col_' cheader]));
                        if sum(contains({'CreatedYear','CreatedMonth','CreatedDay','ShippedYear','ShippedMonth','ShippedDay'},cheader)) == 1
                            value = cell2mat(value);
                        end
                    end
                    
                    % Overwrite address data
                    if sum(contains(tobeoverwritten,cheader)) > 0
                        % Company address data
                        if strcmp(cheader,'HospitalAddress') == 1
                            temp = customers.(cusnr).ClinicAddress.ClinicAddress1;
                            if strcmp(customers.(cusnr).ClinicAddress.ClinicAddress2,'-') == 0
                                temp = [temp ' ' customers.(cusnr).ClinicAddress.ClinicAddress2];
                            end
                            if strcmp(customers.(cusnr).ClinicAddress.ClinicAddress3,'-') == 0
                                temp = [temp ' ' customers.(cusnr).ClinicAddress.ClinicAddress3];
                            end
                            value = temp;
                        end
                        
                        if strcmp(cheader,'HospitalPostalCode') == 1
                            value = customers.(cusnr).ClinicAddress.ClinicPostalCode;
                        end
                        
                        if strcmp(cheader,'HospitalCityName') == 1
                            value = customers.(cusnr).ClinicAddress.ClinicCity;
                        end
                        
                        if strcmp(cheader,'HospitalState') == 1
                            value = customers.(cusnr).ClinicAddress.ClinicState;
                        end
                        
                        if strcmp(cheader,'HospitalCountry') == 1
                            value = customers.(cusnr).ClinicAddress.ClinicCountry;
                        end
                        
                        if strcmp(cheader,'HospitalRegion') == 1
                            if strcmp(upper(customers.(cusnr).ClinicAddress.ClinicCountry),'UNITED STATES') == 1
                                value = 'USA';
                            elseif strcmp(upper(customers.(cusnr).ClinicAddress.ClinicCountry),'CANADA') == 1
                                value = 'CAN';
                            else
                                error(['Hospital country not recognized for customer ' cusnr '. Please check the UPS excel.'])
                            end
                        end
                        
                        
                        
                        % Delivery address data
                        if strcmp(cheader,'DeliveryOfficeAddress') == 1
                            temp = customers.(cshiptonr).DeliveryAddress.DeliveryAddress1;
                            if strcmp(customers.(cshiptonr).DeliveryAddress.DeliveryAddress2,'-') == 0
                                temp = [temp ', ' customers.(cshiptonr).DeliveryAddress.DeliveryAddress2];
                            end
                            if strcmp(customers.(cshiptonr).DeliveryAddress.DeliveryAddress3,'-') == 0
                                temp = [temp ', ' customers.(cshiptonr).DeliveryAddress.DeliveryAddress3];
                            end
                            value = temp;
                        end
                        
                        if strcmp(cheader,'DeliveryPostalCode') == 1
                            value = customers.(cshiptonr).DeliveryAddress.DeliveryPostalCode;
                        end
                        
                        if strcmp(cheader,'DeliveryOfficeCityName') == 1
                            value = customers.(cshiptonr).DeliveryAddress.DeliveryCity;
                        end
                        
                        if strcmp(cheader,'DeliveryOfficeState') == 1
                            value = customers.(cshiptonr).DeliveryAddress.DeliveryState;
                        end
                        
                        if strcmp(cheader,'DeliveryOfficeCountry') == 1
                            value = customers.(cshiptonr).DeliveryAddress.DeliveryCountry;
                        end
                        
                        if strcmp(cheader,'DeliveryOfficeRegion') == 1
                            if strcmp(upper(customers.(cshiptonr).DeliveryAddress.DeliveryCountryCode),'US') == 1
                                value = 'USA';
                            elseif strcmp(upper(customers.(cshiptonr).DeliveryAddress.DeliveryCountryCode),'CA') == 1
                                value = 'CAN';
                            else
                                error(['Delivery office country code not recognized for customer ' cshiptonr '. Please check the UPS excel.'])
                            end
                        end
                        
                    end
                    
                    
                    % Get the data for the initials
                    if strcmp(cheader,'CustomerFirstName') == 1
                        if isa(value{1},'double') == 1
                            firstname_initial = num2str(cell2mat(value));
                        else
                            firstname_initial = char(value);
                            firstname_initial = firstname_initial(1,1);
                        end
                    elseif strcmp(cheader,'CustomerSurname') == 1
                        if isa(value{1},'double') == 1
                            surname_initial = num2str(cell2mat(value));
                        else
                            surname_initial = char(value);
                        end
                        surnamelength = size(char(surname_initial),2);
                        if surnamelength >= 3
                            surname_initial = surname_initial(1,1:3);
                        elseif surnamelength == 2
                            surname_initial = surname_initial(1,1:2);
                        elseif surnamelength == 1
                            surname_initial = surname_initial(1,1);
                        end
                    end
                    
                    if strcmp(cheader,'Initials') == 1
                        % Construct the initials
                        value = [upper(firstname_initial) '. ' upper(surname_initial) '.'];
                    end
                    
                    if strcmp(cheader,'Itemnumber') == 1
                        value = itemnr;
                    end
                    
                    if strcmp(cheader,'InsoleType') == 1
                        currentinsoletype = char(value);
                    end
                    
                    
                    
                    if strcmp(cheader,'Top_size') == 1
                        % Separate the measurement unit from the size number
                        currentsize = char(value);
                        if strcmp(currentsize,'#null') == 1
                            % If no top layer, just update the measturement unit column.
                            % Do nothing
                            measurementunit = 'UK';
                        elseif contains(currentsize,'kids') == 1
                            % For kids sizes
                            currentsize = strrep(currentsize,' kids UK','');
                            % Change the top layer size for when a wide insole type is ordered.
                            if strcmp(currentinsoletype,'Wide') == 1
                                currentsize = str2double(currentsize);
                                currentsize = [num2str(currentsize+1) '.0'];
                            end
                            value = currentsize;
                            % Add a new column at the end with the measurement unit.
                            measurementunit = 'kids UK';
                        else
                            % For adults sizes
                            currentsize = strrep(currentsize,' UK','');
                            % Change the top layer size for when a wide insole type is ordered.
                            if strcmp(currentinsoletype,'Wide') == 1
                                currentsize = str2double(currentsize);
                                currentsize = [num2str(currentsize+1) '.0'];
                            end
                            value = currentsize;
                            % Add a new column at the end with the measurement unit.
                            measurementunit = 'UK';
                        end
                        
                    end
                    
                    if strcmp(cheader,'MeasurementUnit') == 1
                        value = measurementunit;
                    end
                    
                    if strcmp(cheader,'ContactTelephone') == 1
                        value = customers.(cshiptonr).DeliveryAddress.DeliveryPhone;
                    end
                    
                    if strcmp(cheader,'ContactName') == 1
                        value = customers.(cshiptonr).DeliveryAddress.DeliveryContactPerson;
                    end
                    
                    if strcmp(cheader,'ContactEmail') == 1
                        value = customers.(cshiptonr).NotificationEmail;
                    end
                    
                    if strcmp(cheader,'AlternativeCustomerNumber') == 1
                        value = agr_cus.AlternativeCustomerNumber;
                    end
                    
                    
                    
                    % Write the value to file
                    salesordercheck_v3.datafeed(nextentry,ch) = {value};
                    clear value
                end
                clear surname_initial
                clear firstname_initial
                clear currentinsoletype
                clear measurementunit
            end
        end
        % Add the price to the list
        % cc2 = num2str(cc);
        % if cc > 99 && strcmp(cc2(1,end-1:end),'00') == 1
        %     save([values.salesordercheck_folder 'salesordercheck_v3.mat'],'salesordercheck_v3')
        %     if cc > 999  && strcmp(cc2(1,end-2:end),'000') == 1 % && strcmp(values.version,'liveversion') == 1
        %         copyfile([values.salesordercheck_folder 'salesordercheck_v3.mat'],[values.salesordercheck_folder 'temp\Datafeed_' num2str(cc) '_salesordercheck_v3.mat'])
        %     end
        %     disp('File saved');
        % end
    end
end

save([values.salesordercheck_folder '\salesordercheck_v3.mat'],'salesordercheck_v3');
copyfile('Temp\OnlineReport.xlsx',['Output\Navision\' values.y values.mo values.d '_' values.h values.mi values.s '_OnlineReport.xlsx']);
copyfile([values.salesordercheck_folder '\salesordercheck_v3.mat'],['Output\Navision\' values.y values.mo values.d '_' values.h values.mi values.s '_post_salesordercheck_v3.mat']);
disp('File saved');

end