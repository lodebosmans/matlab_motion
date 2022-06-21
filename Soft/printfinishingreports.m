function printfinishingreports()

load Temp\values.mat values
% load Temp\db_production.mat db_production
load Temp\caseids_production.mat caseids_production

% Go over all finishing reports and print them
if values.MenuSelectPrinter > 1
    for x = 1:size(caseids_production,1)  
        
        ccid = strrep(char(caseids_production(x,1)),'-','');
        
        file = [values.finrepfolder  ccid '.xlsx']; % This must be full path name  => only adapt the file in the 'Interface' folder, not in the test environment.
        Excel = actxserver('Excel.Application');
        % Set it to visible
        set(Excel,'Visible',0);
        % Add a Workbook
        Workbooks = Excel.Workbooks;
        Workbook = Workbooks.Open(file);
        %Workbook = invoke(Workbooks, 'Add');
        % Get a handle to Sheets and select Sheet 1
        Sheets = Excel.ActiveWorkBook.Sheets;
        Sheet1 = get(Sheets, 'Item', 1);
        Sheet1.Activate;
        
        % Send to printer
        Excel.ActiveWorkbook.PrintOut(1,1,1,'False',values.MenuSelectPrinter_str);
        
        % Close Excel and clean up
        invoke(Excel,'Quit');
        delete(Excel);
        clear Excel;
    end
end

end