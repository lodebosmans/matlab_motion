function getrushcases()

load Temp\values.mat values
load Temp\caseids_production.mat caseids_production

% Create an output variable
rushcases = cell(0);
% Get the cases to be read in this run
nrofcases_read = size(caseids_production,1);

% Open the rush cases file
[n,t,raw_rush] = xlsread([values.rushcasesfolder values.rushcasesfile],1);
% Copy the file to the fin rep folder
copyfile([values.rushcasesfolder values.rushcasesfile],[values.finrepfolder values.rushcasesfile]);
% Replace the NaN's with a zero
raw_rush(cellfun(@(x) isnumeric(x) && isnan(x), raw_rush)) = {0};
% Get the column indeces
col_caseid = catchcolumnindex({'Case ID'},raw_rush,1);
col_caseid = cell2mat(col_caseid(2,1));
col_comment = catchcolumnindex({'Comment'},raw_rush,1);
col_comment = cell2mat(col_comment(2,1));
col_status = catchcolumnindex({'Status'},raw_rush,1);
col_status = cell2mat(col_status(2,1));
col_prio = catchcolumnindex({'Priority'},raw_rush,1);
col_prio = cell2mat(col_prio(2,1));
col_category = catchcolumnindex({'Category'},raw_rush,1);
col_category = cell2mat(col_category(2,1));
lengthcontent = col_comment-col_caseid+1;

for x = 1:nrofcases_read
    ccid_read = caseids_production(x,1);
    
    % Compare the case IDs with the rush cases file
    [temp] = catchrowindex2(ccid_read,raw_rush,4);
    
    % If there is a match, fetch the data for that specific case ID
    if isempty(temp) == 0
        % Get the length of the temp var, as there can be more than one occurence
        nrofoccur = size(temp,1);
        for coccur = 1:nrofoccur
            % Get the current content of the rushcases variable
            nrofrushes = size(rushcases,1);
            % Get the first following occurence
            nextmatch = temp(coccur,1);
            % Check is the status is on 0 or 1. If 1, ignore case. If 0, proceed.
            if cell2mat(raw_rush(nextmatch,col_status)) == 0
                % Add the data to an overview file
                rushcases(nrofrushes+1,1:lengthcontent) = raw_rush(nextmatch,col_caseid:col_comment);
            end
        end
    else
        % Do nothing
    end
end
rushcases(cellfun(@(x) isnumeric(x), rushcases)) = {'-'};



% List all the open rush cases for and additional overview
rawrushsize = size(raw_rush,1);
rawrushcontent = cell(0);
nrofrushes2 = 0;
for row = 2:rawrushsize
    if cell2mat(raw_rush(row,col_status)) == 0
        nrofrushes2 = nrofrushes2 + 1;
        rawrushcontent(nrofrushes2,1:lengthcontent) = raw_rush(row,col_caseid:col_comment);
    end
end
rawrushcontent(cellfun(@(x) isnumeric(x), rawrushcontent)) = {'-'};

if size(rushcases,1) == 0
    enddoc = 3;
    rowline = enddoc + 1;
    rowbacklogtitle = rowline + 2;
    rowbacklogcontentstart = rowbacklogtitle + 2;
    rowbacklogcontentend = rowbacklogcontentstart + nrofrushes2 - 1;
else
    enddoc = size(rushcases,1) + 2;
    rowline = enddoc + 1;
    rowbacklogtitle = rowline + 2;
    rowbacklogcontentstart = rowbacklogtitle + 2;
    rowbacklogcontentend = rowbacklogcontentstart + nrofrushes2 - 1;
end

enddoc = num2str(enddoc);
rowline = num2str(rowline);
rowbacklogtitle = num2str(rowbacklogtitle);
rowbacklogcontentstart = num2str(rowbacklogcontentstart);
rowbacklogcontentend = num2str(rowbacklogcontentend);

outputwidth = char(xlsColNum2Str(size(rawrushcontent,2)));

% Save the overview file
save Temp\rushcases.mat rushcases

% Print a hard copy of the overview file
Excel = actxserver('Excel.Application');
% Set it to visible
set(Excel,'Visible',1);
% Add a Workbook
Workbooks = Excel.Workbooks;
Workbook = invoke(Workbooks, 'Add');
% Get a handle to Sheets and select Sheet 1
Sheets = Excel.ActiveWorkBook.Sheets;
Sheet1 = get(Sheets, 'Item', 1);
Sheet1.Activate;
% Add the title
eActivesheetRange = get(Excel.Activesheet,'Range','A1:A1');
eActivesheetRange.Value = [values.d '/' values.mo '/20' values.y ' - Rush cases in current shipment - Do not forget to update the status!'];
if nrofrushes2 > 0
    %eActivesheetRange = get(Excel.Activesheet,'Range',['A' num2str(str2double(enddoc) + 1) ':C' num2str(str2double(enddoc) + 1)]);
    eActivesheetRange.Range(['A' rowline ':' outputwidth rowline]).Borders.Item('xlEdgeBottom').LineStyle = 1;
end
% Add the data of the current rush cases
if size(rushcases,1) == 0
    eActivesheetRange = get(Excel.Activesheet,'Range','A3:A3');
    eActivesheetRange.Value = 'No rush cases in the current list of case ID''s.';
else
    eActivesheetRange = get(Excel.Activesheet,'Range',['A3:' outputwidth enddoc]);
    eActivesheetRange.Value = rushcases;
end
% Add some extra text
eActivesheetRange = get(Excel.Activesheet,'Range',['A' rowbacklogtitle ':A' rowbacklogtitle]);
eActivesheetRange.Value = 'Backlog rush cases';
% Add the data of all cases remaining
eActivesheetRange = get(Excel.Activesheet,'Range',['A'  rowbacklogcontentstart ':' outputwidth rowbacklogcontentend]);
eActivesheetRange.Value = rawrushcontent;

% Lay-out the document
% Set the column widths (do not go over 83 in total to keep everything on one page)
Sheet1.Columns.Item(1).columnWidth = 14;
Sheet1.Columns.Item(2).columnWidth = 7;
Sheet1.Columns.Item(3).columnWidth = 9;
Sheet1.Columns.Item(4).columnWidth = 7;
Sheet1.Columns.Item(5).columnWidth = 48;
% Get the column of the comments
colcom = char(xlsColNum2Str(lengthcontent));
colcat = char(xlsColNum2Str(lengthcontent-2));
% And set it to text wrap
eActivesheetRange = get(Excel.Activesheet,'Range',[colcom ':' colcom]);
eActivesheetRange.WrapText = true;
eActivesheetRange = get(Excel.Activesheet,'Range',[colcat ':' colcat]);
eActivesheetRange.WrapText = true;
% Align everything vertically upwards
eActivesheetRange = get(Excel.Activesheet,'Range',['A1:' colcom rowbacklogcontentend]);
eActivesheetRange.VerticalAlignment = 1;

% Create the output filename
outputfile = [pwd '\Output\FinishingReports\' values.y values.mo values.d '_' values.h values.mi values.s '_OverviewRushCases'];

% Calculate the number of pages to print
nrofpages = ceil((6 + size(rawrushcontent,1) + size(rushcases,1))/45); % The 6 is for the fixed amount of lines on the first page.

% Send to printer
if values.MenuSelectPrinter > 1
    Excel.ActiveWorkbook.PrintOut(1,nrofpages,1,'False',values.MenuSelectPrinter_str);
end
% Save as excel
invoke(Workbook,'SaveAs',outputfile)
% Close Excel and clean up
invoke(Excel,'Quit');
delete(Excel);
clear Excel;

end