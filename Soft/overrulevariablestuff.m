function [values] = overrulevariablestuff(values,handles)

[y,mo,d,h,mi,s] = getdate();
values.y = y;
values.mo = mo;
values.d = d;
values.h = h;
values.mi = mi;
values.s = s;

values.mainfolder = 'C:\Users\mathlab\Documents\Matlab\'; % Use full path and include \ at the end
values.baseversion = [values.mainfolder '20190318_Interface_OrFix\']; % Use full path and include \ at the end
values.upsfilename = '20190327_OverviewShipments_v10.xlsx'; 
values.backupfolder = 'O:\Administration\BackupFiles\';
values.backupfoldertestenvironment = 'O:\Administration\BackupFilesTestEnvironment\';
% Can be overwritten
values.upsfilepath = 'O:\StandardOperations\UPS\RS Print\'; % Use full path and include \ at the end -----------------------------------------------------
values.rsfilesfolder = [values.backupfolder 'RSFiles\']; % Use full path and include \ at the end
values.worldshipxmlfilefolder = 'O:\StandardOperations\UPS\RS Print\WorldShip\XmlInput\Input\'; % Use full path and include \ at the end    
values.pollespowerbifolder = 'O:\StandardOperations\Production\FinishingReports\'; % Use full path and include \ at the end

values.TestDebug = get(handles.checkbox_testdebug, 'Value');
values.NoFinRep = get(handles.checkbox_nofinrep, 'Value');

if strcmp(values.version,'liveversion') == 1 && values.TestDebug == 0
    
    values.currentversion = [values.mainfolder '20190318_Interface_OrFix\']; % Use full path and include \ at the end ----------------------------------------------
    values.upsfilepath = 'O:\StandardOperations\UPS\RS Print\'; % Use full path and include \ at the end -----------------------------------------------------
    values.salesordercheck_folder = [values.backupfolder 'Navision\SOC\']; % Use full path and include \ at the end ------------------------------------------------
    values.TestDebug = get(handles.checkbox_testdebug,'Value'); %----------------------------------------------------------------------------------------------
    values.TestDebugstr = '';
    values.navfolder = [values.backupfolder 'Navision\']; % Use full path and include \ at the end
    values.logfolder = [values.backupfolder 'Logfiles\']; % Use full path and include \ at the end
    values.finrepfolder = [values.backupfolder 'FinishingReports\']; % Use full path and include \ at the end
    
elseif strcmp(values.version,'testversion') == 1 || values.TestDebug == 1  

    values.currentversion = [values.mainfolder '20190521_Interface_TestEnvironment_DoNotUse\']; % Use full path and include \ at the end
    values.salesordercheck_folder = [values.backupfoldertestenvironment 'Navision\SOC\']; % Use full path and include \ at the end  
    values.TestDebug = 1;
    values.TestDebugstr = '_TestEnvironment';
    values.navfolder = [values.backupfoldertestenvironment 'Navision\']; % Use full path and include \ at the end
    temp = 'O:\StandardOperations\UPS\RS Print\TestEnvironment\'; % Use full path and include \ at the end
    values.logfolder = [values.backupfoldertestenvironment 'Logfiles\']; % Use full path and include \ at the end
    values.finrepfolder = [values.backupfoldertestenvironment 'FinishingReports\']; % Use full path and include \ at the end
    
    % Copy all essential files to the testenvironment backup folders
    if isfile([values.backupfoldertestenvironment 'TestEnvironmentUpToDate.txt']) == 0
        disp('Copying Navision files to TestEnvironment. Please wait.');
        copyfile([values.backupfolder 'Navision'],[values.backupfoldertestenvironment 'Navision']);
        disp('Copying UPS file to TestEnvironment. Please wait.');
        copyfile([values.upsfilepath values.upsfilename],[temp values.upsfilename]);
        fid2 = fopen([values.backupfoldertestenvironment 'TestEnvironmentUpToDate.txt'],'wt');
        fclose(fid2);
    end
    
    values.upsfilepath = temp;
    values.backupfolder = values.backupfoldertestenvironment;
     
end

% See what needs to be done.

values.onlinereportfolder = 'O:\Administration\BackupFiles\OnlineReport\'; % Use full path and include \ at the end
values.onlinereportfilename = '20200909_v50_RSPrint_OR_new.xlsx'; % With tables included, but not all data
values.onlinereport_tables_copy = [values.onlinereportfolder values.onlinereportfilename]; % Use full path and include \ at the end
values.onlinereportfilename_full = '20200909_v50_RSPrint_OR_Full.xlsx'; % With all data, but no tables
values.onlinereport = [values.onlinereportfolder values.onlinereportfilename]; % Use full path and include \ at the end
values.onlinereportmat = [values.onlinereportfolder 'onlinereport.mat']; % Use full path and include \ at the end
values.consultationfolder = 'O:\StandardOperations\OnlineReport\'; % Use full path and include \ at the end
%values.salesordercheck_file = 'salesordercheck_v3'; % Name is used in too many regulare lines to be replaced..
values.salesordercheckuptodate = 'SOC_v3_UpToDate.txt';
values.agreementsfile = '20190312_OverviewAgreements.xlsx';
values.shipmentexcelfolder = [values.backupfolder 'ShipmentExcels\']; % Use full path and include \ at the end
values.invoicerequired = 0;
values.rushcasesfolder = 'O:\StandardOperations\Production\'; % Use full path and include \ at the end
values.rushcasesfile = 'RushCases.xlsx';
values.invoicesfolder = [values.navfolder 'Invoices\']; % Use full path and include \ at the end
values.deliverynotetemplatefilename = 'DeliveryNoteUpsShipment_v3_format2019.xlsx';
values.deliverynotetemplatefolder = [values.baseversion 'Input\Templates\'];
values.shipmentexceptionsfilename = 'ShipmentExceptions.xlsx';
values.stockboxlabelfilename = 'StockBoxLabelInput.xlsx';
values.stockboxlabelfolder = [values.baseversion 'Input\StockBoxLabel\' values.stockboxlabelfilename];
% Read the production options
values.printercheck = 0;
values.DesiredFunction_nr = get(handles.popupmenu_DesiredFunction,'Value');
values.DesiredFunctions = {'Select an option..','Finishing reports + labels','Delivery note (manual)', ...
                           'Shipment sort out + delivery note + WorldShip xmls + FDA labels','FDA labels', ...
                           'Daily production status snapshot','Get logfile','Check OMS status','Sort shipment Excels','Stock Box Label'};
values.FinRepLabels = 0;
values.ManDelNote = 0;
values.SortOutDelNoteWSxml = 0;
values.FDAlabels = 0;
values.DailySnapshot = 0;
values.OpenLogFile = 0;
values.CheckOMSstatus = 0;
values.LogEvents = 0;
values.SortShipmentExcels = 0;
values.CreateStockBoxLabel = 0;
if values.DesiredFunction_nr == 1
    % Donothing = 1;
elseif values.DesiredFunction_nr == 2
    values.FinRepLabels = 1;
    values.LogEvents = 1;
    values.printercheck = 1;
elseif values.DesiredFunction_nr == 3
    values.ManDelNote = 1;
    values.LogEvents = 1;
    values.printercheck = 1;
elseif values.DesiredFunction_nr == 4
    values.SortOutDelNoteWSxml = 1;
    values.LogEvents = 1;
    values.printercheck = 1;
elseif values.DesiredFunction_nr == 5
    values.FDAlabels = 1;
    values.LogEvents = 1;
elseif values.DesiredFunction_nr == 6
    values.DailySnapshot = 1;
elseif values.DesiredFunction_nr == 7
    values.OpenLogFile = 1;
elseif values.DesiredFunction_nr == 8
    values.CheckOMSstatus = 1;
elseif values.DesiredFunction_nr == 9
    values.SortShipmentExcels = 1;
elseif values.DesiredFunction_nr == 10
    values.CreateStockBoxLabel = 1;
end
% Read the case IDs if present
values.NewCaseID = get(handles.edit_NewCaseID,'String');
if isempty(values.NewCaseID) == 1
    %There is nothing inserted
    values.NewCaseIDsPresent = 0;
else
    values.NewCaseIDsPresent = 1;
    values.NewCaseID = upper(values.NewCaseID);
end
% Read the Navision options
values.NavisionFunction_nr = get(handles.popupmenu_NavisionFunction,'Value');
values.NavisionFunctions = {'Select an option..',...
                            'Generate sales order xmls', ...
                            'Select sales order xmls based on user input',...
                            'Select shipped cases (within one month) to all production facilities + calculate 3D printed pieces',...
                            'Generate sales shipment xmls (shipped within on month) based on UPS Excel (month closing)',...
                            'Generate RS-labels',...
                            'Link invoice to cases/shipments'};
values.NavSalesOrderInput = 0;
values.GrabSalesOrdersUPS = 0;
values.GrabSalesOrdersUser = 0;
values.ListShippedCases = 0;
values.NavSalesShipmentsInput = 0;
values.GenerateRSLabels = 0;
values.LinkInvoiceToCasesShipments = 0;
values.GenerateMonthlyStockCorrections = 0;
values.UpdateOnlineReport = 0;
if values.NavisionFunction_nr == 1
    % Donothing = 1;
elseif values.NavisionFunction_nr == 2
    values.NavSalesOrderInput = 1;
elseif values.NavisionFunction_nr == 3
    values.GrabSalesOrdersUser = 1;
elseif values.NavisionFunction_nr == 4
    values.ListShippedCases = 1;
elseif values.NavisionFunction_nr == 5
    values.NavSalesShipmentsInput = 1;
elseif values.NavisionFunction_nr == 6
    values.GenerateRSLabels = 1;
elseif values.NavisionFunction_nr == 7    
    values.LinkInvoiceToCasesShipments = 1;
elseif values.NavisionFunction_nr == 8    
    values.GenerateMonthlyStockCorrections = 1; 
elseif values.NavisionFunction_nr == 9     
    values.UpdateOnlineReport = 1;
    values.FinRepLabels = 1;
end

if values.NavisionFunction_nr == 2 || values.NavisionFunction_nr == 9 
    % Get the current month/year as well for the OMS status check.
    values.RequestedMonth = values.mo;
    values.RequestedYear = ['20' values.y];
end

list.years = {'Choose a year','2018','2019','2020','2021','2022','2023','2024','2025'};
list.months = {'Choose a month','January','February','March','April','May','June','July','August','September','October','November','December'};
values.MenuSelectPrinter = get(handles.MenuSelectPrinter,'Value');
list.printers = {'Select a printer','NPI50FDAA (HP LaserJet M506)','HPLJ_400'}; 
%When printing errors occur, check which printer is default. This should be 'HP LaserJet M506 PCL 6' on the Matlab PC. Not the network printer (HP LaserJet M506x PCL 6 on RSPRINTDC.RSPRINT.local).
% If no printer is selected, double check if they do not want to select one.
if values.MenuSelectPrinter == 1 && values.printercheck == 1
    [values.MenuSelectPrinter,brol] = listdlg('ListString',list.printers,'SelectionMode','single','PromptString','You have not selected a printer, do you want to select one? If not, press cancel!:','ListSize',[500,100],'Name','Printer selection');  
    if isempty(values.MenuSelectPrinter) == 1
        values.MenuSelectPrinter = 1;
    end

end
values.MenuSelectPrinter_str = char(list.printers(values.MenuSelectPrinter));
list.days = {'Choose a day','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31'};
list.NrOfCopies = {1,2};
values.NrOfCopies = get(handles.CaseidsDeliveryDocument_nrcopies,'Value');
values.NrOfCopies = cell2mat(list.NrOfCopies(values.NrOfCopies));
values.Username_nr = get(handles.MenuUsername,'Value');
list.users = {'Nobody','Julie','Kobe','Lode','Lore','Natascha','Pieter-Jan','Rkia','Sander'};
values.Username_str = char(list.users(values.Username_nr));

save Temp\list.mat list

end