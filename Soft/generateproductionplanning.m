function generateproductionplanning()

load Temp\values.mat values
values.NewCaseID = values.snapshotallcaseids;
save Temp\values.mat values

readcaseids()
sortcaseidsproduction();  
readcasedata()
load Temp\db_production.mat db_production
load Temp\customers.mat customers

col_estshipdate = catchcolumnindex({'estimatedshippingdate'},db_production.headers,1);
col_estshipdate = cell2mat(col_estshipdate(2,1));
col_cusnr = catchcolumnindex({'customernumber'},db_production.headers,1);
col_cusnr = cell2mat(col_cusnr(2,1));
col_caseid = catchcolumnindex({'caseid.full'},db_production.headers,1);
col_caseid = cell2mat(col_caseid(2,1));
col_delcomp = catchcolumnindex({'delivery_company'},db_production.headers,1);
col_delcomp = cell2mat(col_delcomp(2,1));

% Start sorting everything out based on the date
nrofcases = size(db_production.overview,1);
pp = cell(0); % pp = productionplanning
pp(2,1)= {'Shipment date'};
pp(2,2)= {'Counter'};
for cc = 1:nrofcases
    % Get the current size
    [nrr, nrc] = size(pp);
    
    % Get the current case ID
    ccid = char(db_production.overview(cc,col_caseid));
    
    % Get the estimated shipping date
    cdate = db_production.overview(cc,col_estshipdate);
    % Check if the date is already present in the planning
    targetrow = find(strcmp(pp(:,1),cdate));
    if isempty(targetrow) == 1
        % The date is not present yet, so add it to the end.
        targetrow = nrr + 1;
        pp(targetrow,1) = cdate;
    else
        % The date is already present, so it doesn't need to be added.
    end

    % Get the customer number
    ccnr = db_production.overview(cc,col_cusnr);
    ccnr = strrep(ccnr,'-','_');
    ccname = db_production.overview(cc,col_delcomp); %#ok<NASGU>
    % Get the ship to number
    cshipto_nr = customers.(char(ccnr)).ShipTo;
    cshipto_nr = strrep(cshipto_nr,'D','C');
    cshipto_nr = strrep(cshipto_nr,'-','_');
    cshipto_name = customers.(cshipto_nr).Company;
    % Check if the cshipto is already present in the planning
    targetcol = find(strcmp(pp(1,:),cshipto_nr));
    if isempty(targetcol) == 1
        % The customer number is not present yet, so add it to the end.
        targetcol = nrc + 1;
        pp(1,targetcol) = {cshipto_nr};
        pp(2,targetcol) = {cshipto_name};
    else
        % The date is already present, so it doesn't need to be added.
    end
    
    pp(cellfun(@(x) isempty(x), pp)) = {'-'};
    
    % Add the case ID to the target location
    content = [char(pp(targetrow,targetcol)) ' ' ccid];
    content = strrep(content,'- ','');
    pp(targetrow,targetcol) = {content};
    % Increase the counter
    if cellfun(@ischar,pp(targetrow,2)) == 1
        pp(targetrow,2) = {1};
    else
        pp(targetrow,2) = {cell2mat(pp(targetrow,2)) + 1};
    end    
end

pp(1,1:2) = {''};

% Get the current size
[nrr, nrc] = size(pp);
collabel=char(xlsColNum2Str(nrc));

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
eActivesheetRange = get(Excel.Activesheet,'Range',['A1:' collabel num2str(nrr)]);
eActivesheetRange.Value = pp;
eActivesheetRange.WrapText = true;
eActivesheetRange.Font.Size  = 9;
for i = 1:nrc
    eval(['Sheet1.Columns.Item(' num2str(i) ').columnWidth = 11.5']);
end
Sheet1.Columns.Item(2).columnWidth = 7;
eActivesheetRange.VerticalAlignment = 1;
eActivesheetRange = get(Excel.Activesheet,'Range',['A1:A' num2str(nrr)]);
cells = eActivesheetRange.Range('A1:A1000');
set(cells.Font, 'Bold', true)
cells = eActivesheetRange.Range(['A1:' collabel '1']);
set(cells.Font, 'Bold', true)

% And set it to text wrap
eActivesheetRange = get(Excel.Activesheet,'Range','A:C');
eActivesheetRange.WrapText = true;
% Align everything vertically upwards
eActivesheetRange = get(Excel.Activesheet,'Range','A:C');
eActivesheetRange.VerticalAlignment = 1;

Sheet1.PageSetup.Orientation = 1;

Sheet1.PageSetup.LeftMargin = 12.96;  
Sheet1.PageSetup.RightMargin = 36;
Sheet1.PageSetup.TopMargin = 36;
Sheet1.PageSetup.BottomMargin = 36;
Sheet1.PageSetup.HeaderMargin = 36;
Sheet1.PageSetup.FooterMargin = 36;

% Create the output filename
outputfile = [pwd '\Output\Production\' values.y values.mo values.d '_' values.h values.mi values.s '_ProductionPlanning.xlsx'];

% Send to printer
% if values.MenuSelectPrinter > 1
%     Excel.ActiveWorkbook.PrintOut(1,1,1,'False',values.MenuSelectPrinter_str);
% end
% Save as excel
invoke(Workbook,'SaveAs',outputfile);
% Close Excel and clean up
invoke(Excel,'Quit');
delete(Excel);
clear Excel;



%xlswrite('ProductionPlanning',pp)


end



