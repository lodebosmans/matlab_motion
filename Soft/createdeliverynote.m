function createdeliverynote()

load Temp\values.mat values
load Temp\db_production.mat db_production

% Prepare the data
% ----------------
% Get the number of cases
nrofcases = size(db_production.overview,1);
% Check if all from the same company
for x = 1:nrofcases
    if strcmp(char(db_production.overview(1,5)),char(db_production.overview(x,5))) == 1
        donothing = 1;
    elseif strcmp(char(db_production.overview(x,5)),'RS Print') == 1
        donothing = 1;
    else
        msgbox({['The first case in the list for the delivery note is for ' char(db_production.overview(1,5)) ', therefore it is assumed all cases are intended for ' char(db_production.overview(1,5)) '.'],' ',[char(db_production.overview(x,1)) ' is a case for ' char(db_production.overview(x,5)) ' instead of for ' char(db_production.overview(1,5)) '.'],' ','Please verify the content of the delivery.'},'Different companies detected:');
        clear all
    end
end

companyname = char(db_production.overview(1,5));
nrofpages = ceil(nrofcases/60);
nrofcolumns = ceil(nrofcases/30);
lastcolumn = nrofcases - ((nrofcolumns-1)*30);

temp = 1:nrofcases;
temp = temp';
for x = 1:nrofcases
    db_filtered(x,1) = {temp(x,1)};
end
db_filtered(1:nrofcases,2) = db_production.overview(1:nrofcases,1);
for currentcase = 1:nrofcases
    db_filtered(currentcase,3) = {[char(db_production.overview(currentcase,19)) ' ' char(db_production.overview(currentcase,20))]};
end

% Write to file
% -------------
file = [values.baseversion 'Input\Templates\DeliveryNote_v1.xlsx']; % This must be full path name  => only adapt the file in the 'Interface' folder, not in the test environment.
% Open Excel Automation server
Excel = actxserver('Excel.Application');
Workbooks = Excel.Workbooks;
% Make Excel visible
Excel.Visible = 0;
% Open Excel file
Workbook = Workbooks.Open(file);
Sheets = Excel.ActiveWorkBook.Sheets;
sheet1 = get(Sheets, 'Item', 1);
invoke(sheet1, 'Activate');
Activesheet = Excel.Activesheet;
% Add company information
range = 'B4:B4';
data = companyname;
ActivesheetRange = get(Activesheet,'Range',range);
set(ActivesheetRange, 'Value', data);
% Add the page number
range = 'C2:C2';
data = ['Page 1/' num2str(nrofpages)];
ActivesheetRange = get(Activesheet,'Range',range);
set(ActivesheetRange, 'Value', data);

% Some numbers about the formatting of the template
page.rows = 54; % Number of row on the entire page of the template
page.columns = 'F'; % Index of the total amount of colums in the entire page the template
page.startindex = 7; % Relative index of the startposition where the row where the invoice data will be pasted
page.endindex = 36; % Relative index of the endposition where the row where the invoice data will be pasted
page.casespercolumn = (page.endindex - page.startindex + 1); % Number of cases per column in the template
page.casesperpage = 2*page.casespercolumn; % Number of cases per page in the template
%page.indexcolleft = 'A'; % Column where the left numbering will be added
%page.indexcolright = 'D'; % Column where the right numbering will be added

% Copy the file if the nr of pages if larger than 1
range = ['A1:' page.columns num2str(page.rows)];
for extrapage = 1:nrofpages-1
    % Find the start and end position to paste the template
    startrow = (extrapage*page.rows)+1;
    endrow = ((extrapage+1)*page.rows);
    destination = ['A' num2str(startrow) ':' page.columns num2str(endrow)];
    sheet1.Range(range).Copy;
    sheet1.Paste(sheet1.Range(destination));
    % Add the page number
    destination = ['C' num2str(2+(extrapage*page.rows)) ':C' num2str(2+(extrapage*page.rows))];
    ActivesheetRange = get(Activesheet,'Range',destination);
    set(ActivesheetRange, 'Value', ['Page ' num2str(extrapage+1) ' of ' num2str(nrofpages) ]);
    clear startrow
    clear endrow
    clear destination
    clear destination2
    clear destination3 
end

% Insert the data
numbers = 1:nrofcases;
numbers = numbers';
difference = -1;
for currentcolumn = 1:nrofcolumns
    if mod(currentcolumn,2) == 1
        side = 'left';
        col_index = 'A';
        col_case = 'B';
        col_customer = 'C';
    else
        side = 'right';
        col_index = 'D';
        col_case = 'E';
        col_customer = 'F';
    end
    currentpage = ceil(currentcolumn/2);
    % Get the row indices from the raw overview
    input_initialrow = 1 + (currentcolumn-1)*page.casespercolumn;
    input_finalrow = 30 + (currentcolumn-1)*page.casespercolumn;
    if input_finalrow > nrofcases
        difference = input_finalrow - nrofcases;
        input_finalrow = nrofcases;
    end
    output_initialrow = page.startindex + ((currentpage-1)*page.rows);
    output_finalrow = page.endindex + ((currentpage-1)*page.rows);
    if difference > -1
        output_finalrow = output_finalrow - difference;
    end
    % Get the data from the overwiew
    data = db_filtered(input_initialrow:input_finalrow,:);
    % Paste the output in the excel file
    range = [col_index num2str(output_initialrow) ':' col_customer num2str(output_finalrow)];
    ActivesheetRange = get(Activesheet,'Range',range);
    set(ActivesheetRange, 'Value', data);
    
%     numbers = numbers + 30;
end


% % Update the format of the file so everything phits (ba dum tsss) nicely in one page
% sheet1.Columns.Item(1).columnWidth = 3.29;
% sheet1.Columns.Item(2).columnWidth = 10.86;
% sheet1.Columns.Item(3).columnWidth = 18.71;
% sheet1.Columns.Item(4).columnWidth = 9;
% sheet1.Columns.Item(5).columnWidth = 23.43;
% sheet1.Columns.Item(6).columnWidth = 6.29;
% % Set the column width for page 2 and 3
% range = 'N1:AO1';
% sheet1range = Range(sheet1,range);
% sheet1range.ColumnWidth = 8.67;
% % for q = 11:38
% %     eval(['sheet1.Columns.Item(' num2str(q) ').columnWidth = 8.67']);
% % end
% % Set the row height
% range = ['A1:A' num2str(nrdetailpages*page.rows)];
% sheet1range = Range(sheet1,range);
% sheet1range.RowHeight = 11.25;

% Create the output filename
outputfile = [pwd '\Output\DeliveryNotes\' values.y values.mo values.d '_' values.h values.mi values.s '_DeliveryNote_' companyname];
%outputfile2 = [year month day '_' h m s '_DeliveryNote_' companyname];

% Send to printer and print one with names
if values.MenuSelectPrinter > 1
    Excel.ActiveWorkbook.PrintOut(1,nrofpages,1,'False',values.MenuSelectPrinter_str);
end
% Remove the names quick and dirty
ActivesheetRange = get(Activesheet,'Range','C7:C36');
set(ActivesheetRange, 'Value',' ');
ActivesheetRange = get(Activesheet,'Range','F7:F36');
set(ActivesheetRange, 'Value',' ');
% And print one without names
if values.MenuSelectPrinter > 1
    Excel.ActiveWorkbook.PrintOut(1,nrofpages,1,'False',values.MenuSelectPrinter_str);
end

% Save as excel
invoke(Workbook,'SaveAs',outputfile)
% Save as pdf
sheet1.ExportAsFixedFormat('xlTypePDF',outputfile);
% Close Excel and clean up
invoke(Excel,'Quit');
delete(Excel);
clear Excel;

% % Copy the files once more to the server
% destination = 'O:\Administration\DeliveryNotes';
% copyfile([outputfile '.xlsx'],[destination '\' outputfile2 '.xlsx']);
% copyfile([outputfile '.pdf'],[destination '\' outputfile2 '.pdf']);

end