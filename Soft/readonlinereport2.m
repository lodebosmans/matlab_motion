function readonlinereport2()

% Reads the online report and caculates the price of the case. Does not update the itemnumber in the proces
disp('Checking order statusses - updating if necessary - please wait');

load Temp\values.mat values
load Temp\customers.mat customers
load Temp\onlinereport.mat onlinereport
load([values.salesordercheck_folder '\salesordercheck_v3.mat'],'salesordercheck_v3');

headersOR = onlinereport(1,:);
orrows = size(onlinereport,1);
col_year = catchcolumnindex({'CreatedYear'},headersOR,1);
col_year = cell2mat(col_year(2,1));
col_ccid = catchcolumnindex({'CaseCode'},headersOR,1);
col_ccid = cell2mat(col_ccid(2,1));
col_CreatedYear = catchcolumnindex({'CreatedYear'},headersOR,1);
col_CreatedYear = cell2mat(col_CreatedYear(2,1));
col_CreatedMonth = catchcolumnindex({'CreatedMonth'},headersOR,1);
col_CreatedMonth = cell2mat(col_CreatedMonth(2,1));
col_CreatedDay = catchcolumnindex({'CreatedDay'},headersOR,1);
col_CreatedDay = cell2mat(col_CreatedDay(2,1));
col_HospitalName = catchcolumnindex({'HospitalName'},headersOR,1);
col_HospitalName = cell2mat(col_HospitalName(2,1));
col_DeliveryOfficeName = catchcolumnindex({'DeliveryOfficeName'},headersOR,1);
col_DeliveryOfficeName = cell2mat(col_DeliveryOfficeName(2,1));
col_InsoleType = catchcolumnindex({'InsoleType'},headersOR,1);
col_InsoleType = cell2mat(col_InsoleType(2,1));
col_BaseType = catchcolumnindex({'Base_type'},headersOR,1);
col_BaseType = cell2mat(col_BaseType(2,1));
col_CaseCanceled = catchcolumnindex({'CaseCanceled'},headersOR,1);
col_CaseCanceled = cell2mat(col_CaseCanceled(2,1));
col_HospitalName = catchcolumnindex({'HospitalName'},headersOR,1);
col_HospitalName = cell2mat(col_HospitalName(2,1));
col_DeliveryOfficeName = catchcolumnindex({'DeliveryOfficeName'},headersOR,1);
col_DeliveryOfficeName = cell2mat(col_DeliveryOfficeName(2,1));
col_Assembly = catchcolumnindex({'ServiceType'},headersOR,1);
col_Assembly = cell2mat(col_Assembly(2,1));
col_TopMaterial = catchcolumnindex({'Top_material'},headersOR,1);
col_TopMaterial = cell2mat(col_TopMaterial(2,1));

% Sort out all the orders and add the price
% -----------------------------------------

% Get the next line of the salesordercheck variable
cmtlsid = size(salesordercheck_v3.data,1) + 1;
startpoint_OR = cmtlsid + 1; % Compensate for the additional header

if orrows >= startpoint_OR
    % Do the loop for every case.
    for cc = startpoint_OR:orrows
        lineOR = onlinereport(cc,:);
        % Get the case ID
        ccid = char(lineOR(1,col_ccid));
        cmtlsid = cc - 1;
        disp(['Case ' ccid ' - ' num2str(cc) '/' num2str(orrows)]);
        salesordercheck_v3.data(cmtlsid,1) = {ccid};
        if strcmp(lineOR(1,col_CaseCanceled),'-') == 0
            salesordercheck_v3.data(cmtlsid,5) = {'0: Cancelled'};
        else
            if cc == 26572
                test = 1;
            end            
                    
            %         if strcmp(char(lineOR(1,col_HospitalName)),'Ortho Georgia Company') == 1
            %             test = 1;
            %         end
            compname = char(lineOR(1,col_HospitalName));
            delname = char(lineOR(1,col_DeliveryOfficeName));
            disp(['For customer: ' compname ' -  ' delname ]);
            % Get the creation date
            cy = cell2mat(lineOR(1,col_CreatedYear));
            cm = cell2mat(lineOR(1,col_CreatedMonth));
            cd = cell2mat(lineOR(1,col_CreatedDay));
            [itemnr,cusnr] = getitemnumber(ccid,headersOR,customers,lineOR);            
            
            flow = 0;
            nrs.mainnr = '-';
            nrs.agentnr = '-';
            nrs.agentmainnr = '-';
            % Get the correct price agreement based on the creation date
            [agrnr] = getcurrentpriceagreement(customers,cusnr,cy,cm,cd);
            agr_cus = customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr)]);
            % Identify the relevant agreement
            if strcmp(customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr)]).PricingMethod,'Own') == 1
                % If own princing method, use that on.
                flow = 1;
                agr_price = customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr)]);
                nrs.agentnr = customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr)]).Agent;
            elseif strcmp(customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr)]).PricingMethod,'Main') == 1
                % If main account pricing method, get the info from the main account.
                nrs.mainnr = customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr)]).Main;
                [agrnr] = getcurrentpriceagreement(customers,nrs.mainnr,cy,cm,cd);
                if strcmp(customers.(nrs.mainnr).PriceAgreements.(['Agreement' num2str(agrnr)]).PricingMethod,'Own') == 1
                    % If own princing method, use that on.
                    flow = 2;
                    agr_price = customers.(nrs.mainnr).PriceAgreements.(['Agreement' num2str(agrnr)]);
                    nrs.agentnr = customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr)]).Agent;
                elseif strcmp(customers.(nrs.mainnr).PriceAgreements.(['Agreement' num2str(agrnr)]).PricingMethod,'Agent') == 1
                    % If agent account pricing method, get the info from the agent account.
                    nrs.agentnr = customers.(nrs.mainnr).PriceAgreements.(['Agreement' num2str(agrnr)]).Agent;
                    [agrnr] = getcurrentpriceagreement(customers,nrs.agentnr,cy,cm,cd);
                    if strcmp(customers.(nrs.agentnr).PriceAgreements.(['Agreement' num2str(agrnr)]).PricingMethod,'Own') == 1
                        % If own princing method, use that on.
                        flow = 3;
                        agr_price = customers.(nrs.agentnr).PriceAgreements.(['Agreement' num2str(agrnr)]);
                    elseif strcmp(customers.(nrs.agentnr).PriceAgreements.(['Agreement' num2str(agrnr)]).PricingMethod,'Main') == 1
                        % If main agent account pricing method, get the info from the main account.
                        flow = 4;
                        nrs.agentmainnr = customers.(nrs.agentnr).PriceAgreements.(['Agreement' num2str(agrnr)]).Main;
                        [agrnr] = getcurrentpriceagreement(customers,nrs.agentmainnr,cy,cm,cd);
                        agr_price = customers.(nrs.agentmainnr).PriceAgreements.(['Agreement' num2str(agrnr)]);
                    else
                        disp('Fatal error!')
                        clear all
                    end
                end
            elseif strcmp(customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr)]).PricingMethod,'Agent') == 1
                % If agent account pricing method, get the info from the agent account.
                nrs.agentnr = customers.(cusnr).PriceAgreements.(['Agreement' num2str(agrnr)]).Agent;
                [agrnr] = getcurrentpriceagreement(customers,nrs.agentnr,cy,cm,cd);
                if strcmp(customers.(nrs.agentnr).PriceAgreements.(['Agreement' num2str(agrnr)]).PricingMethod,'Own') == 1
                    % If own princing method, use that on.
                    flow = 5;
                    agr_price = customers.(nrs.agentnr).PriceAgreements.(['Agreement' num2str(agrnr)]);
                elseif strcmp(customers.(nrs.agentnr).PriceAgreements.(['Agreement' num2str(agrnr)]).PricingMethod,'Main') == 1
                    % If main agent account pricing method, get the info from the main account.
                    flow = 6;
                    nrs.agentmainnr = customers.(nrs.agentnr).PriceAgreements.(['Agreement' num2str(agrnr)]).Main;
                    [agrnr] = getcurrentpriceagreement(customers,nrs.agentmainnr,cy,cm,cd);
                    agr_price = customers.(nrs.agentmainnr).PriceAgreements.(['Agreement' num2str(agrnr)]);
                else
                    disp('Fatal error!')
                    clear all
                end
            else
                disp('Fatal error!')
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
                material = char(lineOR(1,col_TopMaterial));
                assembly = char(lineOR(1,col_Assembly));
                basetype = char(lineOR(1,col_BaseType));
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
            

            % Add the price to the list
            cc2 = num2str(cc);
            if strcmp(cc2(1,end-1:end),'00') == 1
                save([values.salesordercheck_folder 'salesordercheck_v3.mat'],'salesordercheck_v3')
                disp('File saved');
            end
        end
    end
end

save([values.salesordercheck_folder '\salesordercheck_v3.mat'],'salesordercheck_v3');
disp('File saved');

end