function printfootscandeclaration()

load Temp\values.mat values

% Specify file name
file = [values.baseversion 'Input\Templates\DeclarationOfConformityFootscan_v2.xlsx']; % This must be full path name
% Open Excel Automation server
Excel = actxserver('Excel.Application');
Workbooks = Excel.Workbooks;
% Make Excel visible or not
Excel.Visible = 0;
% Open Excel file
Workbook = Workbooks.Open(file); %#ok<NASGU>
Excel.ActiveWorkbook.PrintOut(1,1,1,'False','NPI50FDAA (HP LaserJet M506)');
% Close Excel and clean up
invoke(Excel,'Quit');
delete(Excel);
clear Excel;

end