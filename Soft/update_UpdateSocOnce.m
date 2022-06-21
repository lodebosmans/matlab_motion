% A file to update the SOC once with several columns
% Starting position is the SOC with 17 colums

% Load the necessary files
load Temp\values.mat values
load Temp\customers.mat customers
load Temp\onlinereport.mat onlinereport
load([values.salesordercheck_folder 'salesordercheck_v3.mat'],'salesordercheck_v3');

% {'CaseID','CustomerNumber','CompanyName','DeliveryFacility','Status','ItemNumber','CreationDate','Price','PriceComment','LocationProduction','ShipmentDate','ShipmentService','TrackingNumber','ReferenceRSPrint','ReshipmentInfo','InvoiceNumber','FinishingReportPrint','ShipmentCost','MaterialCost','LaborCost','Currency','MarketSegment','Country','PurchaseCost3dPart'}

% Define the new colums
newcolumns = {'ShipmentCost','MaterialCost','LaborCost','Currency','MarketSegment','Country','PurchaseCost3dPart','PurchaseCostAssembledItem'};

% Add the colums
for x = 1:size(newcolumns,2)
    newcol = newcolumns(1,x);
    % Check if the column is not already present
    if sum(contains(salesordercheck_v3.headers,newcol)) == 0
        % Add the column header
        salesordercheck_v3.headers(1,end+1) = newcol;
        % Add the column with dashed lines
        salesordercheck_v3.data(:,end+1) = {'-'};
    end
end

col_status = catchcolumnindex({'Status'},salesordercheck_v3.headers,1);
col_status = cell2mat(col_status(2,1));
col_currency = catchcolumnindex({'Currency'},salesordercheck_v3.headers,1);
col_currency = cell2mat(col_currency(2,1));
col_marketsegment = catchcolumnindex({'MarketSegment'},salesordercheck_v3.headers,1);
col_marketsegment = cell2mat(col_marketsegment(2,1));
col_country = catchcolumnindex({'Country'},salesordercheck_v3.headers,1);
col_country = cell2mat(col_country(2,1));
col_shipmentcost = catchcolumnindex({'ShipmentCost'},salesordercheck_v3.headers,1);
col_shipmentcost = cell2mat(col_shipmentcost(2,1));
col_customernumber = catchcolumnindex({'CustomerNumber'},salesordercheck_v3.headers,1);
col_customernumber = cell2mat(col_customernumber(2,1));

% Go over the orders and fill that table
for cr = 1:size(salesordercheck_v3.data,1)
    disp(cr);
    if cr == 21000
        test = 1;
    end
    % Get the status
    % Continue if not cancelled or pre-navision era
    if strcmp(char(salesordercheck_v3.data(cr,col_status)),'0: Cancelled') == 0 && strcmp(char(salesordercheck_v3.data(cr,col_status)),'Pre-Navision Era') == 0
        % Get the customer number
        cnr = char(salesordercheck_v3.data(cr,col_customernumber));
        cnr_ = strrep(cnr,'-','_');
        
        % Update the currency
        
        % Update the market segment
        salesordercheck_v3.data(cr,col_marketsegment) = {customers.(cnr_).MarketSegment};
        
        % Update the country
        salesordercheck_v3.data(cr,col_country) = {customers.(cnr_).InvoiceAddress.InvoiceCountry};
        
        % Update the shipment cost
        
    end
    
end