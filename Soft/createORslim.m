function createORslim()

load Temp\values.mat values

disp('Trimming down the OR for the big boss');

% Create the filename
ORslim_filename = [values.y values.mo values.d '_' values.h values.mi values.s '_OnlineReport_VoorDePolle.xlsx'];
ORslim_path = [values.consultationfolder ORslim_filename];

% ORslim_path = [values.consultationfolder 'TestFile.xlsx'];

% Copy the file
copyfile([values.onlinereport],ORslim_path);
% copyfile('TestFile.xlsx',ORslim_path);


[type, sheet_names] = xlsfinfo(ORslim_path);
% First open Excel as a COM Automation server
Excel = actxserver('Excel.Application'); 
% Make the application invisible
set(Excel, 'Visible', 0);
% Make excel not display alerts
set(Excel,'DisplayAlerts',0);
% Get a handle to Excel's Workbooks
Workbooks = Excel.Workbooks; 
% Open an Excel Workbook and activate it
Workbook=Workbooks.Open(ORslim_path);
% Get the sheets in the active Workbook
Sheets = Excel.ActiveWorkBook.Sheets;
index_adjust = 0;
% Cycle through the sheets and delete them based on user input
for i = 2:max(size(sheet_names))
    current_sheet = get(Sheets, 'Item', (i-index_adjust));
    invoke(current_sheet, 'Delete')
    out_string = sprintf('Worksheet called %s deleted',sheet_names{i});
    index_adjust = index_adjust +1;
    disp(out_string);
    disp(' ');
end
% Now save the workbook
Workbook.Save;
% Close the workbook
Workbooks.Close;
% Quit Excel
invoke(Excel, 'Quit');
% Delete the handle to the ActiveX Object
delete(Excel);

clc

end