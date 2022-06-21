function createnavisionsalesorderinput()

% Goes over the cases to check if an sales order can be created. Fetches price from the global overview (salesordercheck.data).

% Code to renew the sales order check
%         salesordercheck_v3 = cell(100000,3);
%         salesordercheck_v3(1:100000,1) = {0};
%         salesordercheck_v3(1:100000,2:3) = {'Empty'};
%         save 'C:\Users\mathlab\Documents\MATLAB\Interface\Input\salesordercheck_v3.mat' salesordercheck_v3

%load Temp\UPSfile_shipment.mat UPSfile_shipment %#ok<NASGU>
load Temp\values.mat values
load Temp\customers customers
load Temp\onlinereport.mat onlinereport
if isfile('Temp\salesordercheck_v3.mat') == 0
    copyfile([values.salesordercheck_folder 'salesordercheck_v3.mat'],'Temp\salesordercheck_v3.mat'); % always keep full pathway
end
load Temp\salesordercheck_v3.mat salesordercheck_v3
load Temp\UPSfile_customer.mat UPSfile_customer

disp('Creating Navision sales order input - please wait')

% Open file to write all sales orders that are processed in this run
fid = fopen(['Output\Navision\' values.y values.mo values.d '_' values.h values.mi values.s '_SO_XMLsPrepared.txt'],'wt');
% Take backup of salesordercheck_v3 variable.
copyfile([values.salesordercheck_folder 'salesordercheck_v3.mat'],['Output\Navision\' values.y values.mo values.d '_' values.h values.mi values.s '_pre_salesordercheck_v3.mat']);

headersOR = onlinereport(1,:);
col_caseid_mtls = catchcolumnindex({'CaseId'},headersOR,1);
col_caseid_mtls = cell2mat(col_caseid_mtls(2,1));
col_caseid_motion = catchcolumnindex({'CaseCode'},headersOR,1);
col_caseid_motion = cell2mat(col_caseid_motion(2,1));
col_builtactor = catchcolumnindex({'BuiltActor'},headersOR,1);
col_builtactor = cell2mat(col_builtactor(2,1));
% col_shippedactor = catchcolumnindex({'ActorShipped'},headersOR,1);
% col_shippedactor = cell2mat(col_shippedactor(2,1));
col_casecancelled = catchcolumnindex({'CaseCanceled'},headersOR,1);
col_casecancelled = cell2mat(col_casecancelled(2,1));
col_hospitalname = catchcolumnindex({'HospitalName'},headersOR,1);
col_hospitalname = cell2mat(col_hospitalname(2,1));
col_deliveryofficename = catchcolumnindex({'DeliveryOfficeName'},headersOR,1);
col_deliveryofficename = cell2mat(col_deliveryofficename(2,1));
col_insoletype = catchcolumnindex({'InsoleType'},headersOR,1);
col_insoletype = cell2mat(col_insoletype(2,1));
col_createdyear = catchcolumnindex({'CreatedYear'},headersOR,1);
col_createdyear = cell2mat(col_createdyear(2,1));
col_shippedmonth = catchcolumnindex({'ShippedMonth'},headersOR,1);
col_shippedmonth = cell2mat(col_shippedmonth(2,1));
col_topmaterial = catchcolumnindex({'Top_material'},headersOR,1);
col_topmaterial = cell2mat(col_topmaterial(2,1));

headersSOC = salesordercheck_v3.headers(1,:);
col_price = catchcolumnindex({'Price'},headersSOC,1);
col_price = cell2mat(col_price(2,1));

[OR_endpoint_row,nrofcols] = size(onlinereport); %#ok<ASGLU>

OR_endpoint_ID = cell2mat(onlinereport(end,col_caseid_mtls));

OR_startpoint_row = ceil(OR_endpoint_row*0.80);

for cr = OR_startpoint_row:OR_endpoint_row
    if cr == 105707
        test = 1;
    end
    OR_current_ID = cell2mat(onlinereport(cr,col_caseid_mtls));
    OR_current_caseID = char(onlinereport(cr,col_caseid_motion));
    disp(['Reading line ' num2str(OR_current_ID) '/' num2str(OR_endpoint_ID) ' of OnlineReport: ' OR_current_caseID]);
    %         if strcmp(caseidrsp,'RS18-PEX-PUH') == 1
    %             test = 1;
    %         end
    %     if cr == 13556
    %         test = 1;
    %     end
   
    
    % Check if the the sales order has been made already.
    cstatus = char(salesordercheck_v3.data(OR_current_ID,5));
    disp(cstatus);
    if strcmp(cstatus(1),'0') == 1 || ...
            strcmp(cstatus,'0: Cancelled') == 1 || strcmp(cstatus,'0: Pre-Navision Era') == 1 || ...
            strcmp(cstatus,'Cancelled') == 1 || strcmp(cstatus,'Pre-Navision Era') == 1 || ...
            strcmp(cstatus,'Empty') == 1
        % Skip the case, because it has already been processed.
        % skip = 1;
        disp(' ');
        
    else
        % Check if in status 1 (temporary itemnumber) or 2 (final itemnumber)
        if strcmp(cstatus(1),'2') == 1  || strcmp(cstatus(1),'1') == 1
            
            % Check if itemnumber has progressed from 1 to 2.
            [itemnr,brol] = getitemnumber(OR_current_caseID,headersOR,customers,UPSfile_customer,onlinereport(cr,:));
            
            % If still temporary or no built actor yet, skip this case.
            temp_ba = catchcolumnindex({'BuiltActor'},headersOR,1);
            temp_ba = char(onlinereport(cr,cell2mat(temp_ba(2,1))));
            temp_b = catchcolumnindex({'Built'},headersOR,1);
            temp_b = char(onlinereport(cr,cell2mat(temp_b(2,1))));
            
            % Location is defined here, but als in getitemnumber line 104 => double check this!
            if strcmp('MPROD US Materialise',temp_ba) == 1 ...
                    || strcmp('Flowbuilt Production',temp_ba) == 1 ...
                    || strcmp('Flowbuilt Production_temp',temp_ba) == 1
                location = 'USPROD';
            else
                location = 'BEPROD';
            end
            
            if strcmp(itemnr(2),'X') == 0 && strcmp('-',temp_ba) == 0 && strcmp('-',temp_b) == 0
                disp('Not temporary anymore, progressed to final itemnumber.')
                cusdata.company = char(onlinereport(cr,col_hospitalname));
                cusdata.delivery = char(onlinereport(cr,col_deliveryofficename));
                
                if strcmp(cusdata.company,'RS Print Company') == 1 ...
                        || strcmp(cusdata.company,'RS Print Marketing BE Company') == 1 ...
                        || strcmp(cusdata.company,'RS Print Marketing US Company') == 1 ...
                        || strcmp(cusdata.company,'RS Print Marketing CBE Company') == 1 ...
                        || strcmp(cusdata.company,'RS Print Marketing CUS Company') == 1 ...
                        || strcmp(cusdata.company,'RS Print Sponsoring Company') == 1 ...
                        || strcmp(cusdata.company,'RS Print Production BE Company') == 1 ...
                        || strcmp(cusdata.company,'RS Print Production US Company') == 1 ...
                        || strcmp(cusdata.company,'RS Print Support BE Company') == 1 ...
                        || strcmp(cusdata.company,'RS Print Support US Company') == 1 ...
                        || strcmp(cusdata.company,'RS Print Reorder BE Company') == 1 ...
                        || strcmp(cusdata.company,'RS Print Reorder US Company') == 1 ...
                        || strcmp(cusdata.company,'RS Print Development BE Company') == 1
                    RSP = 1;
                else
                    RSP = 0;
                end
                [cnr] = getcustomernumber(cusdata,UPSfile_customer);
                message = ['Processing sales order ID ' num2str(OR_current_ID) '/' num2str(OR_endpoint_ID) ' (' OR_current_caseID ') for ' cusdata.company ' - ' cusdata.delivery ' (' cnr ')'];
                fprintf(fid,message);
                fprintf(fid, '\n');
                disp(message);
                cnr = strrep(cnr,'-','_');
                %             % Catch row index of current case
                %             rowccOR = catchrowindex({caseidrsp},onlinereport,2);
                %             rowccOR = cell2mat(rowccOR(2,1));
                
                % Write to file
                
                docNode = com.mathworks.xml.XMLUtils.createDocument('SalesOrder');
                docRootNode = docNode.getDocumentElement;
                
                % ------------------------------------------------------------- %
                %                        SalesHeader                            %
                % ------------------------------------------------------------- %
                
                SalesHeader_tag = docNode.createElement('SalesHeader');
                docRootNode.appendChild(SalesHeader_tag);
                
                temp = customers.(cnr).CustomerNumber;
                SellToCustomerNo_tag = docNode.createElement('SellToCustomerNo');
                SellToCustomerNo_text = docNode.createTextNode(temp);
                SellToCustomerNo_tag.appendChild(SellToCustomerNo_text);
                SalesHeader_tag.appendChild(SellToCustomerNo_tag);
                
                temp = catchcolumnindex({'ReferenceID'},headersOR,1);
                temp = onlinereport(cr,cell2mat(temp(2,1)));
                if isa(temp{1},'double') == 1
                    temp = num2str(cell2mat(temp));
                else
                    temp = char(temp);
                end
                CustomerReference_tag = docNode.createElement('CustomerReference');
                CustomerReference_text = docNode.createTextNode(temp);
                CustomerReference_tag.appendChild(CustomerReference_text);
                SalesHeader_tag.appendChild(CustomerReference_tag);
                
                temp2 = catchcolumnindex({'SurgeryDate'},headersOR,1);
                temp2 = char(onlinereport(cr,cell2mat(temp2(2,1))));
                indexes = strfind(temp2,'/');
                first = indexes(1,1);
                second = indexes(1,2);
                dateday = temp2(1,1:first-1);
                if length(dateday) == 1
                    dateday = ['0' dateday];
                end
                datemonth = temp2(1,first+1:second-1);
                if length(datemonth) == 1
                    datemonth = ['0' datemonth];
                end
                dateyear = temp2(1,second+1:end);
                %                 if length(temp2) == 10
                %                     % donothing = 1;
                %                 else
                %                     temp2 = [ '0' temp2 ];
                %                 end
                % Reverse order to format YYYY/MM/DD
                temp2 = [ dateyear '/' datemonth '/' dateday ];
                RequestedDeliveryDate_tag = docNode.createElement('RequestedDeliveryDate');
                RequestedDeliveryDate_text = docNode.createTextNode(temp2);
                RequestedDeliveryDate_tag.appendChild(RequestedDeliveryDate_text);
                SalesHeader_tag.appendChild(RequestedDeliveryDate_tag);
                
                temp2 = catchcolumnindex({'CreatedYear'},headersOR,1);
                createdyear = num2str(cell2mat(onlinereport(cr,cell2mat(temp2(2,1)))));
                temp2 = catchcolumnindex({'CreatedMonth'},headersOR,1);
                temp2 = cell2mat(onlinereport(cr,cell2mat(temp2(2,1))));
                if temp2 < 10
                    createdmonth = ['0' num2str(temp2)];
                else
                    createdmonth = num2str(temp2);
                end
                temp2 = catchcolumnindex({'CreatedDay'},headersOR,1);
                temp2 = cell2mat(onlinereport(cr,cell2mat(temp2(2,1))));
                if temp2 < 10
                    createdday = ['0' num2str(temp2)];
                else
                    createdday = num2str(temp2);
                end
                temp = [createdyear '/' createdmonth '/' createdday ];
                OrderDate_tag = docNode.createElement('OrderDate');
                OrderDate_text = docNode.createTextNode(temp);
                OrderDate_tag.appendChild(OrderDate_text);
                SalesHeader_tag.appendChild(OrderDate_tag);
                
                CaseID_tag = docNode.createElement('CaseID');
                CaseID_text = docNode.createTextNode(OR_current_caseID);
                CaseID_tag.appendChild(CaseID_text);
                SalesHeader_tag.appendChild(CaseID_tag);
                
                SalesPersonCode_tag = docNode.createElement('SalesPersonCode');
                SalesPersonCode_text = docNode.createTextNode('');
                SalesPersonCode_tag.appendChild(SalesPersonCode_text);
                SalesHeader_tag.appendChild(SalesPersonCode_tag);
                
                Location_tag = docNode.createElement('Location');
                Location_text = docNode.createTextNode(location);
                Location_tag.appendChild(Location_text);
                SalesHeader_tag.appendChild(Location_tag);
                
                temp = customers.(cnr).ShipTo;
                ShipToCode_tag = docNode.createElement('ShipToCode');
                ShipToCode_text = docNode.createTextNode(temp);
                ShipToCode_tag.appendChild(ShipToCode_text);
                SalesHeader_tag.appendChild(ShipToCode_tag);
                
                ShipmentMethodCode_tag = docNode.createElement('ShipmentMethodCode');
                ShipmentMethodCode_text = docNode.createTextNode('');
                ShipmentMethodCode_tag.appendChild(ShipmentMethodCode_text);
                SalesHeader_tag.appendChild(ShipmentMethodCode_tag);
                
                ShippingAgentCode_tag = docNode.createElement('ShippingAgentCode');
                ShippingAgentCode_text = docNode.createTextNode('');
                ShippingAgentCode_tag.appendChild(ShippingAgentCode_text);
                SalesHeader_tag.appendChild(ShippingAgentCode_tag);
                
                ShippingAgentServiceCode_tag = docNode.createElement('ShippingAgentServiceCode');
                ShippingAgentServiceCode_text = docNode.createTextNode('');
                ShippingAgentServiceCode_tag.appendChild(ShippingAgentServiceCode_text);
                SalesHeader_tag.appendChild(ShippingAgentServiceCode_tag);
                
                % ------------------------------------------------------------- %
                %                         SalesLine                             %
                % ------------------------------------------------------------- %
                
                SalesLine_tag = docNode.createElement('SalesLine');
                SalesHeader_tag.appendChild(SalesLine_tag);
                
                Type_tag = docNode.createElement('Type');
                Type_text = docNode.createTextNode('Item');
                Type_tag.appendChild(Type_text);
                SalesLine_tag.appendChild(Type_tag);
                
                [itemnr,brol] = getitemnumber(OR_current_caseID,headersOR,customers,UPSfile_customer,onlinereport(cr,:));
                ItemNo_tag = docNode.createElement('ItemNo');
                ItemNo_text = docNode.createTextNode(itemnr);
                ItemNo_tag.appendChild(ItemNo_text);
                SalesLine_tag.appendChild(ItemNo_tag);
                
                PurchasingCode_tag = docNode.createElement('PurchasingCode');
                PurchasingCode_text = docNode.createTextNode('');
                PurchasingCode_tag.appendChild(PurchasingCode_text);
                SalesLine_tag.appendChild(PurchasingCode_tag);
                
                description = char(onlinereport(cr,col_insoletype));
                % If Livit, add the top layer material
                if strcmp(cnr(2:5),'1071')
                    description = [description ' - ' char(onlinereport(cr,col_topmaterial))];
                end
                Description2_tag = docNode.createElement('Description2');
                %Description2_text = docNode.createTextNode(description);
                % Add the difference between standard and slim
                if strcmp(itemnr(2:3),'10') == 1
                    Description2_text = docNode.createTextNode(['BE semi regular - ' description]);
                elseif strcmp(itemnr(2:3),'11') == 1
                    Description2_text = docNode.createTextNode(['BE full regular - ' description]);
                elseif strcmp(itemnr(2:3),'12') == 1
                    Description2_text = docNode.createTextNode(['BE semi slim - ' description]);
                elseif strcmp(itemnr(2:3),'13') == 1
                    Description2_text = docNode.createTextNode(['BE full slim - ' description]);
                elseif strcmp(itemnr(2:3),'20') == 1
                    Description2_text = docNode.createTextNode(['US semi regular - ' description]);
                elseif strcmp(itemnr(2:3),'21') == 1
                    Description2_text = docNode.createTextNode(['US full regular - ' description]);
                elseif strcmp(itemnr(2:3),'22') == 1
                    Description2_text = docNode.createTextNode(['US semi slim - ' description]);
                elseif strcmp(itemnr(2:3),'23') == 1
                    Description2_text = docNode.createTextNode(['US full slim - ' description]);
                else
                    clear all
                    error('The insole type is not known.')
                end
                % Overrule if no toplayer itemnr
                if strcmp(itemnr(15),'N') == 1
                    if strcmp(itemnr(2:3),'10') == 1
                        Description2_text = docNode.createTextNode('BE semi regular - No top layer');
                    elseif strcmp(itemnr(2:3),'11') == 1
                        Description2_text = docNode.createTextNode('BE full regular - No top layer');
                    elseif strcmp(itemnr(2:3),'12') == 1
                        Description2_text = docNode.createTextNode('BE semi slim - No top layer');
                    elseif strcmp(itemnr(2:3),'13') == 1
                        Description2_text = docNode.createTextNode('BE full slim - No top layer');
                    elseif strcmp(itemnr(2:3),'20') == 1
                        Description2_text = docNode.createTextNode('US semi regular - No top layer');
                    elseif strcmp(itemnr(2:3),'21') == 1
                        Description2_text = docNode.createTextNode('US full regular - No top layer');
                    elseif strcmp(itemnr(2:3),'22') == 1
                        Description2_text = docNode.createTextNode('US semi slim - No top layer');
                    elseif strcmp(itemnr(2:3),'23') == 1
                        Description2_text = docNode.createTextNode('US full slim - No top layer');
                    else
                        clear all
                    end
                end
                Description2_tag.appendChild(Description2_text);
                SalesLine_tag.appendChild(Description2_tag);
                
                LocationCode_tag = docNode.createElement('LocationCode');
                LocationCode_text = docNode.createTextNode(location);
                LocationCode_tag.appendChild(LocationCode_text);
                SalesLine_tag.appendChild(LocationCode_tag);
                
                Quantity_tag = docNode.createElement('Quantity');
                Quantity_text = docNode.createTextNode('1');
                Quantity_tag.appendChild(Quantity_text);
                SalesLine_tag.appendChild(Quantity_tag);
                
                UnitOfMeasureCode_tag = docNode.createElement('UnitOfMeasureCode');
                UnitOfMeasureCode_text = docNode.createTextNode('');
                UnitOfMeasureCode_tag.appendChild(UnitOfMeasureCode_text);
                SalesLine_tag.appendChild(UnitOfMeasureCode_tag);
                
                % Get the price for this case.
                currentprice = num2str(cell2mat(salesordercheck_v3.data(OR_current_ID,col_price)));
                UnitPriceExclVAT_tag = docNode.createElement('UnitPriceExclVAT');
                UnitPriceExclVAT_text = docNode.createTextNode(currentprice);
                UnitPriceExclVAT_tag.appendChild(UnitPriceExclVAT_text);
                SalesLine_tag.appendChild(UnitPriceExclVAT_tag);
                
                LineDiscount_tag = docNode.createElement('LineDiscount');
                if RSP == 1
                    LineDiscount_text = docNode.createTextNode('100');
                else
                    LineDiscount_text = docNode.createTextNode('');
                end
                LineDiscount_tag.appendChild(LineDiscount_text);
                SalesLine_tag.appendChild(LineDiscount_tag);
                clear RSP
                
                LineDiscountAmount_tag = docNode.createElement('LineDiscountAmount');
                LineDiscountAmount_text = docNode.createTextNode('');
                LineDiscountAmount_tag.appendChild(LineDiscountAmount_text);
                SalesLine_tag.appendChild(LineDiscountAmount_tag);
                
                
                % ------------------------------------------------------------- %
                %                       ItemTrackingLine                        %
                % ------------------------------------------------------------- %
                
                ItemTrackingLine_tag = docNode.createElement('ItemTrackingLine');
                SalesLine_tag.appendChild(ItemTrackingLine_tag);
                
                SerialNo_tag = docNode.createElement('SerialNo');
                SerialNo_text = docNode.createTextNode(OR_current_caseID);
                SerialNo_tag.appendChild(SerialNo_text);
                ItemTrackingLine_tag.appendChild(SerialNo_tag);
                
                % ------------------------------------------------------------- %
                %                          Wrap up                              %
                % ------------------------------------------------------------- %
                
                % docNode.appendChild(docNode.createComment('this is a comment'));
                xmlFileName = [values.navfolder 'SalesOrders\All\SalesOrder_' num2str(OR_current_ID) '_' OR_current_caseID '.xml'];
                xmlwrite(xmlFileName,docNode);
                %type(xmlFileName);
                
                salesordercheck_v3.data(OR_current_ID,5) = {'3: Sales order xml created'};
                salesordercheck_v3.data(OR_current_ID,6) = {itemnr};
            else
                disp('Skipped for now: either still temporary itemnumber or not yet passed the production phase.')
            end
        else
            % Do nothing, this case is already futher in the process.
        end
        disp(' ');
        
        % cr2 = num2str(cr);
        % if strcmp(cr2(1,end-2:end),'000') == 1 % Dont change to 00, because it will pause every 100 records to save the file.
        %     save([values.salesordercheck_folder '\salesordercheck_v3.mat'],'salesordercheck_v3')
        %     copyfile([values.salesordercheck_folder 'salesordercheck_v3.mat'],[values.salesordercheck_folder 'temp\XMfiles_' cr2 '_salesordercheck_v3.mat'])
        %     disp('File saved');
        %     disp(' ');
        % end
    end
end

% save the file (in regular interface folder - not in test folder)
save Temp\salesordercheck_v3.mat salesordercheck_v3 % Save in the temp folder
save([values.salesordercheck_folder '\salesordercheck_v3.mat'],'salesordercheck_v3') % Save on the final destination
%disp(' ');
disp('File saved');
disp(' ');
copyfile([values.salesordercheck_folder '\salesordercheck_v3.mat'],['Output\Navision\' values.y values.mo values.d '_' values.h values.mi values.s '_post_salesordercheck_v3.mat']);
copyfile('Temp\OnlineReport.xlsx',['Output\Navision\' values.y values.mo values.d '_' values.h values.mi values.s '_OnlineReport.xlsx']);

fclose(fid);

% % Create the file that indicates that the salesordercheck_v3 variable is up to date.
% fid2 = fopen([values.onlinereportfolder 'SOC_v3_UpToDate.txt'],'wt');
% fclose(fid2);
% Copy the online report .mat file to the online report folder
% copyfile('Temp\onlinereport.mat',values.onlinereportmat); % MOVED to readonlinereport.m to speed up the process.

end