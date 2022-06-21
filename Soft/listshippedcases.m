function listshippedcases()

load Temp\values.mat values
copyfile([values.salesordercheck_folder '\salesordercheck_v3.mat'],'Temp\salesordercheck_v3.mat'); % always keep full pathway
load Temp\salesordercheck_v3.mat salesordercheck_v3

disp(['Summarizing the shipped cases for ' values.RequestedMonth '/' values.RequestedYear ' - please wait']);

col_itemnumber = catchcolumnindex({'ItemNumber'},salesordercheck_v3.headers,1);
col_itemnumber = cell2mat(col_itemnumber(2,1));

content = 0;
outputfile = [pwd '\Output\Navision\' values.RequestedYear values.RequestedMonth '_ShippedCases.xlsx'];

output_US = cell(0);
output_BE = cell(0);

% Get all folders
mainfolders = dir([values.shipmentexcelfolder 'Sorted']);
nrofmainfolders = length(mainfolders);

% Go over main every folder
for mf = 1:nrofmainfolders
    mainfoldername = mainfolders(mf).name;
    if strcmp(mainfoldername,'US') == 1 || strcmp(mainfoldername,'FB_DEALER') == 1
        target = 'output_US';
    else
        target = 'output_BE';
    end
    % Only go over useful folders
    if strcmp(mainfoldername,'.') == 0 && strcmp(mainfoldername,'..') == 0
        % Check if the year/month folder exists
        foldername_full = [values.shipmentexcelfolder 'Sorted\' mainfoldername '\' values.RequestedYear '\' values.RequestedMonth];
        if isfolder(foldername_full) == 1
            % Get the amount of Excel files in this folder
            subfiles = dir(foldername_full);
            nrofsubfiles = length(subfiles);
            if nrofsubfiles > 2
                for sf = 1:nrofsubfiles
                    subfilename = subfiles(sf).name;
                    if strcmp(subfilename,'.') == 0 && strcmp(subfilename,'..') == 0
                        % Double check if the file(s) is/are in the correct folder
                        if strcmp([values.RequestedYear values.RequestedMonth],subfilename(1:6)) == 1
                            % Read the content of the file
                            fullpathfilename = [foldername_full '\' subfilename];
                            disp(['Reading ' fullpathfilename ' - please wait']);
                            clear a1
                            clear a2
                            clear tmp
                            clear tmp2
                            clear raw
                            clear textraw
                            [a1,textraw,raw] = xlsread(fullpathfilename);
                            
                            
                            if contains(fullpathfilename,'.csv')
                                % Convert for the FB csv files.
                                if size(textraw,1) == 1
                                    tempxx = split(textraw,',')';
                                else
                                    tempxx = split(textraw,',');
                                end
                                textraw = tempxx(2:end,1);
                            end
                            nrofcases_temp = size(textraw,1);
                            %nrofcases_temp = size(raw,1);
                            if nrofcases_temp > 0
                                nrofcases_final = 0;
                                % Check if the amount is correct (and doesn't contain dummy cells) + add the 3D printed part
                                for tmp = 1:nrofcases_temp
                                    ccid = char(textraw(tmp,1));
                                    ccid = strrep(ccid,' ','');
                                    if isempty(ccid) == 0
                                        nrofcases_final = nrofcases_final + 1;
                                    end
                                    % Find the itemnumber in de salesordercheck
                                    disp(ccid);
                                    tmp2 = catchrowindex({ccid},salesordercheck_v3.data,1);
                                    tmp2 = cell2mat(tmp2(2,1));
                                    itemnr = char(salesordercheck_v3.data(tmp2,col_itemnumber));
                                    if strcmp('X',itemnr(2)) == 1
                                        if strcmp(target,'output_BE') == 1
                                            d3part(tmp,1) = {['3D1' itemnr(3)]};
                                        elseif strcmp(target,'output_US') == 1
                                            d3part(tmp,1) = {['3D2' itemnr(3)]};
                                        else
                                            error('Dit een een nutteloze error boodschap, maar er is wel iets mis.');
                                        end
                                        %d3part(tmp,1) = {'unknown'};
                                    else
                                        d3part(tmp,1) = {['3D' itemnr(2:3)]};
                                    end
                                end
                                % Get the size of the output file and add to the end
                                eval(['startpoint = size(' target ',1) + 1;']);
                                endpoint = startpoint + nrofcases_final - 1;
                                eval([target '(' num2str(startpoint) ':' num2str(endpoint) ',1) = textraw(1:' num2str(nrofcases_final) ',1);']);
                                eval([target '(' num2str(startpoint) ':' num2str(endpoint) ',2) = d3part;']);
                                eval([target '(' num2str(startpoint) ',3) = {mainfoldername};']);
                                eval([target '(' num2str(startpoint) ',4) = {' num2str(nrofcases_final) '};']);
                                eval([target '(' num2str(startpoint) ',5) = {fullpathfilename};']);
                                
                                content = content + 1;
                                clear d3part
                            end
                        else
                            disp(['File ' subfilename ' might be incorrectly located in ' foldername_full '. Please check.']);
                        end
                    end
                end
            else
                % Nothing to report for mainfoldername
            end
        end
    end
end

% [colChar] = xlsColNum2Str(csize_col);
% colChar = char(colChar);

% Write all data to an Excel file
% Get handle to Excel COM Server
Excel = actxserver('Excel.Application');
% Set it to visible
set(Excel,'Visible',0);
% Add a Workbook
Workbooks = Excel.Workbooks;
Workbook = invoke(Workbooks, 'Add');
% Get a handle to Sheets and select Sheet 1
Sheets = Excel.ActiveWorkBook.Sheets;

targets = {'output_BE','output_US'};
countries = {'Belgium','United States'};
nroftargets = size(targets,2);
for ctarget = 1:nroftargets
    target = char(targets(1,ctarget));
    country  = char(countries(1,ctarget));
    % Count the total amount of 3D printed parts
    csize = eval(['size(' target ',1);']);
    csize_col = eval(['size(' target ',2);']);
    part3D10 = 0;
    part3D11 = 0;
    part3D12 = 0;
    part3D13 = 0;
    part3D20 = 0;
    part3D21 = 0;
    part3D22 = 0;
    part3D23 = 0;
    partunknown = 0;
    for crow = 1:csize
        ctype = eval(['char(' target '(' num2str(crow) ',2));']);
        ctype = ctype(3:4);
        if strcmp(ctype,'10') == 1
            part3D10 = part3D10 + 2;
        elseif strcmp(ctype,'11') == 1
            part3D11 = part3D11 + 2;
        elseif strcmp(ctype,'12') == 1
            part3D12 = part3D12 + 2;
        elseif strcmp(ctype,'13') == 1
            part3D13 = part3D13 + 2;
        elseif strcmp(ctype,'20') == 1
            part3D20 = part3D20 + 2;
        elseif strcmp(ctype,'21') == 1
            part3D21 = part3D21 + 2;
        elseif strcmp(ctype,'22') == 1
            part3D22 = part3D22 + 2;
        elseif strcmp(ctype,'23') == 1
            part3D23 = part3D23 + 2;
        elseif strcmp(ctype,'kn') == 1
            partunknown = partunknown + 2;
        else
            error('Part not known.');
        end
    end
    
    strings = {'3D10','3D11','3D12','3D13','3D20','3D21','3D22','3D23','unknown'};
    sizestrings = size(strings,2);
    for y = 1:sizestrings
        cstring = strings(1,y);
        eval([target '(csize + 2 + ' num2str(y) ',1) = cstring;']);
        eval([target '(csize + 2 + ' num2str(y) ',2) = {part' char(cstring) '};']);
    end
    
    [colChar] = xlsColNum2Str(csize_col);
    if strcmp(colChar,'@') == 1
        colChar = 'B';
    else
        eval(['colChar_' target ' = char(colChar);']);
    end
    
    if ctarget == 1
        eval(['Sheet' num2str(ctarget) ' = get(Sheets, ''Item'', 1);']);
    else
        % Sheet is added in the beginning of the worksheet
        eval(['Sheet' num2str(ctarget) ' = get(Sheets, ''add'');']);
    end
    eval(['Sheet' num2str(ctarget) '.Activate;']);
    % Inject the data
    eval(['size_current = size(' target ',1);']);
    eActivesheetRange = get(Excel.Activesheet,'Range',['A1:' char(colChar) num2str(size_current)]);
    eval(['eActivesheetRange.Value = ' target ';']);
    eval(['Sheet' num2str(ctarget) '.Columns.Item(1).columnWidth = 15;']);
    % Rename the first tab in the worksheet
    Workbook.Worksheets.Item(1).Name = country;
    
end


% Save file
invoke(Workbook, 'SaveAs', outputfile);
invoke(Excel, 'Quit');

end