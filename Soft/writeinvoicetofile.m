function [values] =  writeinvoicetofile(invoicedetails,values)

% Calculate total cost
invoicedetails.totalcost = sum(cell2mat(invoicedetails.overview(:,12)));
% Calculate the number of pages
nrdetailpages = ceil(invoicedetails.counter/30);
% Specify file name
%file = 'C:\Users\mathlab\Documents\MATLAB\Interface\Input\Templates\Invoice6_US.xlsx';
if strcmp(invoicedetails.currency,'USD')
    file = 'C:\Users\mathlab\Documents\MATLAB\Interface\Input\Templates\Invoice7_US.xlsx'; % This must be full path name  => only adapt the file in the 'Interface' folder, not in the test environment.
else
    file = 'C:\Users\mathlab\Documents\MATLAB\Interface\Input\Templates\Invoice7_BE.xlsx';
end
% Open Excel Automation server
Excel = actxserver('Excel.Application');
Workbooks = Excel.Workbooks;
% Make Excel visible
Excel.Visible=1;
% Open Excel file
Workbook = Workbooks.Open(file);
Sheets = Excel.ActiveWorkBook.Sheets;
sheet1 = get(Sheets, 'Item', 1);
invoke(sheet1, 'Activate');
Activesheet = Excel.Activesheet;
% Write the currency in the template
range = 'K8:K8;M8:M8';
data = ['Price (' invoicedetails.currency  ')' ];
ActivesheetRange = get(Activesheet,'Range',range);
set(ActivesheetRange, 'Value', data);
% Insert customer name
range = 'B4:B4';
data = invoicedetails.fullname;
ActivesheetRange = get(Activesheet,'Range',range);
set(ActivesheetRange, 'Value', data);

% Insert date
range = 'B5:B5';
if values.InvoiceDay_nr < 10
    day = ['0' num2str(values.InvoiceDay_nr)];
else 
    day = num2str(values.InvoiceDay_str);
end
if values.InvoiceMonth_nr < 10
    month = ['0' num2str(values.InvoiceMonth_nr)];
else
    month = num2str(values.InvoiceMonth_nr);
end
data = [day '/' values.InvoiceMonth_str '/' values.InvoiceYear_str];
ActivesheetRange = get(Activesheet,'Range',range);
set(ActivesheetRange, 'Value', data);
% Insert invoice number (first to 'ifs' are identical - try to combine both)
if (values.MenuYear + values.MenuMonth >= 4 && invoicedetails.invoicemonthly == 1) || (values.MenuYear + values.MenuMonth == 2 && invoicedetails.invoicealternatively == 1)
    [invoicenumber] = getinvoicenumber(values);
    data = ['S' num2str(values.InvoiceYear_nr) month '-' invoicenumber];
    values.NextInvoiceNumber = values.NextInvoiceNumber + 1;
else
    data = 'NoInvoiceNumber';
end
invoicenumber = data;
range = 'B6:B6'; % 'B6:B6;B52:B52';
ActivesheetRange = get(Activesheet,'Range',range);
set(ActivesheetRange, 'Value', data);
% Copy first page
range = 'A1:M46';
destination = 'A47:M92';
sheet1.Range(range).Copy;
sheet1.Paste(sheet1.Range(destination));

% % Insert date
% range = 'B5:B5';
% data = date;
% ActivesheetRange = get(Activesheet,'Range',range);
% set(ActivesheetRange, 'Value', data);
% % Copy first page
% range = 'A1:M46';
% destination = 'A47:M92';
% sheet1.Range(range).Copy;
% sheet1.Paste(sheet1.Range(destination));
% % Insert invoice number
% if values.MenuYear + values.MenuMonth >= 4 && invoicedetails.invoicemonthly == 1
%     [invoicenumber] = getinvoicenumber(values);
%     data = ['VKF17' invoicenumber];
%     values.NextInvoiceNumber = values.NextInvoiceNumber + 1;
% elseif values.MenuYear + values.MenuMonth == 2 && invoicedetails.invoicealternatively == 1
%     [invoicenumber] = getinvoicenumber(values);
%     data = ['VKF17' invoicenumber];
%     values.NextInvoiceNumber = values.NextInvoiceNumber + 1;
% else
%     data = 'NoInvoiceNumber';
% end
% invoicenumber = data;
% range = 'B6:B6;B52:B52';
% ActivesheetRange = get(Activesheet,'Range',range);
% set(ActivesheetRange, 'Value', data);

% Some numbers about the formatting of the template
page.rows = 46; % Number of row on the entire page of the template
page.columns = 'M'; % Index of the total amount of colums in the entire page the template
page.totalrow = 40; % Row of the position of the word total 
page.totalcol = 'K'; % Column of the position of the word total
page.startindex = 10; % Relative index of the startposition where the row where the invoice data will be pasted
page.endindex = 39; % Relative index of the endposition where the row where the invoice data will be pasted
page.casesperpage = page.endindex - page.startindex + 1; % Number of cases per page in the template
page.indexcol = 'A'; % Column where the numbering will be added
page.totalnumberofpages = nrdetailpages + 3; % Includes the sales conditions and signature
% Add some first page numbers -- (only completely valid if there is only one page of invoice details (i.e. max 30 orders)). Will be overruled later on if there is more than 1 page of invoice details.
destination = ['A' num2str(page.rows) ':M' num2str(page.rows)];
ActivesheetRange = get(Activesheet,'Range',destination);
set(ActivesheetRange, 'Value', [invoicenumber ' - Page 1 of ' num2str(page.totalnumberofpages) ]);
destination = ['A' num2str(page.rows*2) ':M' num2str(page.rows*2)];
ActivesheetRange = get(Activesheet,'Range',destination);
set(ActivesheetRange, 'Value', [invoicenumber ' - Page 2 of ' num2str(page.totalnumberofpages) ]);
destination = ['N' num2str(page.rows) ':AA' num2str(page.rows)];
ActivesheetRange = get(Activesheet,'Range',destination);
set(ActivesheetRange, 'Value', [invoicenumber ' - Page ' num2str(page.totalnumberofpages-1) ' of ' num2str(page.totalnumberofpages) ]);
destination = ['N' num2str(page.rows*2) ':AA' num2str(page.rows*2)];
ActivesheetRange = get(Activesheet,'Range',destination);
set(ActivesheetRange, 'Value', [invoicenumber ' - Page ' num2str(page.totalnumberofpages) ' of ' num2str(page.totalnumberofpages) ]);
% Copy the number of pages required if more than one page
if nrdetailpages > 1
    range = ['A1:' page.columns num2str(page.rows)];
    for extrapage = 1:nrdetailpages
        % Find the start and end position to paste the template
        startrow = (extrapage*page.rows)+1;
        endrow = ((extrapage+1)*page.rows);
        destination = ['A' num2str(startrow) ':' page.columns num2str(endrow)];
        sheet1.Range(range).Copy;
        sheet1.Paste(sheet1.Range(destination));
        % Add the page number
        destination = ['A' num2str(endrow) ':M' num2str(endrow)];
        ActivesheetRange = get(Activesheet,'Range',destination);
        set(ActivesheetRange, 'Value', [invoicenumber ' - Page ' num2str(extrapage+1) ' of ' num2str(page.totalnumberofpages) ]);
        % Add the 'total' string on the last page
        if extrapage == nrdetailpages-1
            % Add the currency
            rowtotal = (extrapage*page.rows)+ page.totalrow;
            destination2 = [page.totalcol num2str(rowtotal) ':' page.totalcol num2str(rowtotal)];
            ActivesheetRange = get(Activesheet,'Range',destination2);
            set(ActivesheetRange, 'Value', ['Total (' invoicedetails.currency  ')']);
            % Add cost
            destination3 = [page.columns num2str(rowtotal) ':' page.columns num2str(rowtotal)];
            ActivesheetRange = get(Activesheet,'Range',destination3);
            set(ActivesheetRange, 'Value', num2str(invoicedetails.totalcost));
        end
        clear startrow
        clear endrow
        clear destination
        clear destination2
        clear destination3
    end
    clear x
    % Copy the signature as last page of the second column to avoid empty pages
%     range = 'AB1:AO46';
%     destination = 'N47:AA92';
%     sheet1.Range(range).Cut;
%     sheet1.Paste(sheet1.Range(destination));
    %clear range
    %clear destination
else
    ActivesheetRange = get(Activesheet,'Range',[page.totalcol num2str(page.totalrow) ':' page.totalcol num2str(page.totalrow)]);
    set(ActivesheetRange, 'Value', ['Total (' invoicedetails.currency  ')']);
    % add cost
    rowtotal = page.totalrow;
    destination3 = [page.columns num2str(rowtotal) ':' page.columns num2str(rowtotal)];
    ActivesheetRange = get(Activesheet,'Range',destination3);
    set(ActivesheetRange, 'Value', num2str(invoicedetails.totalcost));
end
% Insert the data
numbers = 1:page.casesperpage;
numbers = numbers';
for pagenr = 1:nrdetailpages
    % Get the row indices from the raw overview
    overview.initialrow = ((pagenr-1)*page.casesperpage) + 1;
    if pagenr == nrdetailpages 
        overview.finalrow = size(invoicedetails.overview,1);
    elseif pagenr < nrdetailpages
        overview.finalrow = (pagenr*page.casesperpage);
    end
    % Get the row indices for the output file
    output.initialrow = ((pagenr-1)*page.rows) + page.startindex;
    if pagenr == nrdetailpages 
        output.finalrow = output.initialrow + (page.casesperpage - ((nrdetailpages*page.casesperpage) - invoicedetails.counter)) - 1;
    elseif pagenr < nrdetailpages        
        output.finalrow = ((pagenr-1)*page.rows) + page.endindex;
    end
    % Get the data from the overwiew
    data = invoicedetails.overview(overview.initialrow:overview.finalrow,:);
    % Paste the data in the output file
    range = ['B' num2str(output.initialrow ) ':' page.columns num2str(output.finalrow)];
    ActivesheetRange = get(Activesheet,'Range',range);
    set(ActivesheetRange, 'Value', data);
    % Add the index numbers
    if pagenr > 1
        numbers = numbers + 30;
    end
    range = [page.indexcol num2str(output.initialrow ) ':' page.indexcol num2str(output.finalrow)];
    ActivesheetRange = get(Activesheet,'Range',range);
    set(ActivesheetRange, 'Value', numbers);  
end
% Add summary data on last detail page
headertop = (nrdetailpages * page.rows) + page.startindex-2;
headerbottom = (nrdetailpages * page.rows) + page.startindex-1;
range = ['A' num2str(headertop) ':H' num2str(headerbottom)];
sheet1.Range(range).ClearContents;
sheet1.Range(['B' num2str(headertop)]).Value = 'Summary';
sheet1.Range(['C' num2str(headertop)]).Value = 'Insole type';
sheet1.Range(['F' num2str(headertop)]).Value = 'Amount';
sheet1.Range(['G' num2str(headertop)]).Value = 'Price/piece';
fieldnameslist = fieldnames(invoicedetails.summary);
position = headerbottom + 1;
for ct = 1:length(fieldnameslist)
    currenttype = invoicedetails.summary.(char(fieldnameslist(ct,1))).name;
    currentamount = invoicedetails.summary.(char(fieldnameslist(ct,1))).counter;
    currentvat = invoicedetails.summary.(char(fieldnameslist(ct,1))).vat;
    currentprice = invoicedetails.summary.(char(fieldnameslist(ct,1))).price;
    
    sheet1.Range(['C' num2str(position)]).Value = currenttype;
    sheet1.Range(['F' num2str(position)]).Value = currentamount;
    sheet1.Range(['G' num2str(position)]).Value = currentprice;
    sheet1.Range(['K' num2str(position)]).Value = (currentamount * currentprice);
    sheet1.Range(['L' num2str(position)]).Value = currentvat;
    sheet1.Range(['M' num2str(position)]).Value = (currentamount * currentprice) * (1 + (currentvat/100));
    position = position + 1;
end
% Update the format of the file so everything phits (ba dum tsss) nicely in one page
sheet1.Columns.Item(1).columnWidth = 3.29;
sheet1.Columns.Item(2).columnWidth = 10.86;
sheet1.Columns.Item(3).columnWidth = 18.71;
sheet1.Columns.Item(4).columnWidth = 9;
sheet1.Columns.Item(5).columnWidth = 23.43;
sheet1.Columns.Item(6).columnWidth = 6.29;
sheet1.Columns.Item(7).columnWidth = 11.29;
sheet1.Columns.Item(8).columnWidth = 7.86;
sheet1.Columns.Item(9).columnWidth = 6.43;
sheet1.Columns.Item(10).columnWidth = 6.43;
sheet1.Columns.Item(11).columnWidth = 6.86;
sheet1.Columns.Item(12).columnWidth = 3.29;
sheet1.Columns.Item(13).columnWidth = 6.86;
% Set the column width for page 2 and 3
range = 'N1:AO1';
sheet1range = Range(sheet1,range);
sheet1range.ColumnWidth = 8.67;
% for q = 11:38
%     eval(['sheet1.Columns.Item(' num2str(q) ').columnWidth = 8.67']);
% end
% Set the row height
range = ['A1:A' num2str(nrdetailpages*page.rows)];
sheet1range = Range(sheet1,range);
sheet1range.RowHeight = 11.25;
% for q = 1:nrofpages*page.rows
%     eval(['sheet1.Rows.Item(' num2str(q) ').rowHeight = 11.25']);
% end

% % Restore range for correct pdf creation
% if nrofpages == 1
%     range = ['A1:AL' num2str(nrofpages*page.rows)];
%     sheet1 = Range(sheet1,range);
% else
%     
% end

% if isempty(values.ManualInvoice) == 1
%     outputfile = [pwd '\Output\Invoices\' invoicenumber '_' invoicedetails.fullname2];
% else
    outputfile = [pwd '\Output\Invoices\' invoicenumber '_' invoicedetails.fullname2];
% end
% Save as excel
invoke(Workbook,'SaveAs',outputfile)
% Save as pdf
sheet1.ExportAsFixedFormat('xlTypePDF',outputfile);
% Close Excel and clean up
invoke(Excel,'Quit');
delete(Excel);
clear Excel;



end