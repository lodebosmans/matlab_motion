function varargout = InterfaceTest4(varargin)
% InterfaceTest4 MATLAB code for InterfaceTest4.fig
%      InterfaceTest4, by itself, creates a new InterfaceTest4 or raises the existing
%      singleton*.
%
%      H = InterfaceTest4 returns the handle to a new InterfaceTest4 or the handle to
%      the existing singleton*.
%H
%      InterfaceTest4('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in InterfaceTest4.M with the given input arguments.
%
%      InterfaceTest4('Property','Value',...) creates a new InterfaceTest4 or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before InterfaceTest4_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to InterfaceTest4_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help InterfaceTest4

% Last Modified by GUIDE v2.5 29-Jul-2021 17:02:55

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
    'gui_Singleton',  gui_Singleton, ...
    'gui_OpeningFcn', @InterfaceTest4_OpeningFcn, ...
    'gui_OutputFcn',  @InterfaceTest4_OutputFcn, ...
    'gui_LayoutFcn',  [] , ...
    'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before InterfaceTest4 is made visible.
function InterfaceTest4_OpeningFcn(hObject, eventdata, handles, varargin) %#ok<*INUSL>
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to InterfaceTest4 (see VARARGIN)

% Choose default command line output for InterfaceTest4
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes InterfaceTest4 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = InterfaceTest4_OutputFcn(hObject, eventdata, handles)
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;

% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Delete all files from the previous run
clearfolders();

values = struct;
if contains(pwd,'TestEnvironment') == 0
    values.version = 'liveversion';
else
    values.version = 'testversion';
end

[values] = overrulevariablestuff(values,handles);

% ---------------------------------------------------------------------------------------------------------------------

[values] = warningmessages(values);

% If we are going for the daily production snapshot, ask for the case IDs
if values.DailySnapshot == 1
  [values] = getdailysnapshot(values);    
end

% If we are going for a monthly wrap up of information, ask for the year/month of interest
if values.GrabSalesOrdersUPS == 1 || values.ListShippedCases == 1 || values.NavSalesShipmentsInput  == 1 ...
        || values.CheckOMSstatus == 1 || values.GenerateRSLabels == 1 || values.GenerateMonthlyStockCorrections == 1
    if values.GrabSalesOrdersUPS == 1
        [values] = askforyearmonth(values,'SO');
    end
    if values.NavSalesShipmentsInput == 1
        [values] = askforyearmonth(values,'SS');
    end
    if values.ListShippedCases == 1
        [values] = askforyearmonth(values,'ShippedCases');
    end
    if values.CheckOMSstatus == 1
        [values] = askforyearmonth(values,'CheckOMSStatus');
    end
    if values.GenerateRSLabels == 1
        [values] = askforyearmonth(values,'GenerateRSLabels');
    end
    if values.GenerateMonthlyStockCorrections == 1
        [values] = askforyearmonth(values,'StockCorrections');
    end
end

% Save the values variable
save Temp\values.mat values
SOC_backup = [values.backupfolder 'MatlabValues\' values.y values.mo values.d '_' values.h values.mi values.s '_' values.Username_str values.TestDebugstr];
mkdir(SOC_backup);
eval(['save ' SOC_backup '\' values.y values.mo values.d '_' values.h values.mi values.s '_' values.Username_str '_values.mat values']);

% Sort the shipment excel files
if values.SortShipmentExcels == 1
    sortshipmentexcels()
end

% Before reading all heavy files, check if all rs-files are present for the shipments.
readcaseids();
if values.SortOutDelNoteWSxml == 1 && values.NewCaseIDsPresent == 1
    checkforrsnumbers();
end

% Log all the actions that will occur: STEP 1 - initiation.
if (values.LogEvents == 1 && values.NewCaseIDsPresent == 1) || values.DailySnapshot == 1
    logevents('Initiation');
end

tic
if values.OpenLogFile == 0
    % Read the UPS file (first tab with customer information)
    if values.FinRepLabels == 1 || values.ManDelNote == 1 || values.SortOutDelNoteWSxml == 1 ...
            || values.FDAlabels == 1 || values.NavSalesOrderInput == 1 || values.CheckOMSstatus == 1 ...
            || isfile([values.onlinereportfolder values.salesordercheckuptodate]) == 0 ...
            || values.DailySnapshot == 1 || values.UpdateOnlineReport == 1 ...
            || values.NavSalesShipmentsInput
        
        readUPSfile(1);
    end
    % Read the other UPS file tabs if necessary
    if values.FinRepLabels == 1 || values.ManDelNote == 1  || values.SortOutDelNoteWSxml == 1 || values.NavSalesOrderInput == 1 || values.FDAlabels == 1 ...
            || values.CheckOMSstatus == 1 ...
            || values.NavSalesShipmentsInput == 1 ...
            || isfile([values.onlinereportfolder values.salesordercheckuptodate]) == 0 ...
            || values.DailySnapshot == 1 ...
            || values.UpdateOnlineReport == 1 ...
            || values.NavSalesShipmentsInput
        
        if values.DailySnapshot == 0
            readUPSfile(2); % Shipments
            readUPSfile(4); % Pricing
        end
        if values.SortOutDelNoteWSxml == 1
            readUPSfile(5); % Itemnumbers
        end
%         if values.CheckOMSstatus == 0 % You don't need it for the OMS status check.
        readagreements(); % Custom agreements_v2
%         end
    end
    if values.FinRepLabels == 1 || values.ManDelNote == 1 || values.SortOutDelNoteWSxml == 1 || values.NavSalesOrderInput == 1 || values.FDAlabels == 1 ...
            || isfile([values.onlinereportfolder values.salesordercheckuptodate]) == 0 || values.DailySnapshot == 1 || values.UpdateOnlineReport == 1 ...
            || values.NavSalesShipmentsInput
        
        createcustomeroverview();
    end
end
toc

% Read the OnlineReport if necessary
if values.invoicerequired == 1 || values.FinRepLabels == 1 ...
        || values.ManDelNote == 1 || values.FDAlabels == 1 || values.CheckOMSstatus == 1 ...
        || values.SortOutDelNoteWSxml == 1 ...
        || isfile([values.onlinereportfolder values.salesordercheckuptodate]) == 0 || values.NavSalesOrderInput == 1 ...
        || values.NavSalesShipmentsInput == 1 || values.ManDelNote == 1 ...
        || values.DailySnapshot == 1 ...
        || values.GenerateMonthlyStockCorrections == 1 ...
        || values.UpdateOnlineReport == 1
    
    %if isfile([values.onlinereportfolder values.salesordercheckuptodate]) == 1
    if isfile(values.onlinereportmat) == 1 % && isfile([values.onlinereportfolder 'OR_go4d.mat']) && isfile([values.onlinereportfolder 'OR_flowbuilt.mat'])
        disp('Copying online report from database.');
        copyfile(values.onlinereportmat,'Temp\onlinereport.mat');
        copyfile(values.onlinereport,'Temp\OnlineReport.xlsx');
    else
        readonlinereport();
    end
end

% Sort the case IDs that have been read from the input field
if isfile('Temp\caseids_production.mat') == 1 && (values.FinRepLabels == 1 || values.SortOutDelNoteWSxml == 1 || values.FDAlabels == 1 || values.ManDelNote == 1)
    sortcaseidsproduction();
end

% Sort out all the orders of the past for invoicing, add price and determine itemnumber + generate XML files
% Can be ignored if there is an error in the process
if strcmp(values.CountryOfShipment,'BE') == 1
    % if values.IgnoreNavXmls == 0 && isfile([values.onlinereportfolder values.salesordercheckuptodate]) == 0 && values.NavSalesOrderInput == 1
    if isfile([values.onlinereportfolder values.salesordercheckuptodate]) == 0 && (values.NavSalesOrderInput == 1 || values.UpdateOnlineReport == 1)
        try
            copyfile([values.onlinereport_tables_copy],[values.consultationfolder values.y values.mo values.d '_' values.h values.mi values.s '_' values.onlinereportfilename]);
        catch
        end
%         try
%             copyfile([values.onlinereport_tables_copy],[values.consultationfolder values.y values.mo values.d '_' values.h values.mi values.s '_OnlineReport_VoorTom.xlsx']);
%         catch
%         end
        try
            copyfile([values.onlinereport_tables_copy],[values.consultationfolder values.y values.mo values.d '_' values.h values.mi values.s '_OnlineReport_VoorNatascha.xlsx']);
        catch
        end
        try
            copyfile([values.onlinereport_tables_copy],[values.consultationfolder values.y values.mo values.d '_' values.h values.mi values.s '_OnlineReport_VoorIedereen.xlsx']);
        catch
        end
%         try
%             copyfile([values.onlinereport],[values.consultationfolder values.y values.mo values.d '_' values.h values.mi values.s '_OnlineReport_VoorMichael.xlsx']);
%         catch
%         end
        if values.UpdateOnlineReport == 1
            readonlinereport3();
        end
        if values.NavSalesOrderInput == 1 || values.UpdateOnlineReport == 1
            createnavisionsalesorderinput();
        end
        % Save as the SOC as excel file
        load([values.salesordercheck_folder '\salesordercheck_v3.mat'],'salesordercheck_v3') 
        temp_SOC(1,:) = salesordercheck_v3.headers;
        temp_SOC(2:size(salesordercheck_v3.data,1)+1,:) = salesordercheck_v3.data(:,:);
        xlswrite([values.salesordercheck_folder '\salesordercheck_v3.xlsx'],temp_SOC);
        % Create the file that indicates that the salesordercheck_v3 variable is up to date.
        fid2 = fopen([values.onlinereportfolder 'SOC_v3_UpToDate.txt'],'wt');
        fclose(fid2);
        
        senddatatorsscanportal();
        checkomsstatus();
    end
end

% Create labels and finishing reports
if (values.FinRepLabels == 1 || values.UpdateOnlineReport == 1) && values.NoFinRep == 0
    % Get the case IDs
    %readcaseids();
    if isfile('Temp\caseids_production.mat')
        % Read the case data of the selected cases
        readcasedata();
        % Create the labels
        createlabels();
        % Check for rush cases
        try
            getrushcases();
        catch
            disp('The rush cases file is not available.');
            error('...');
        end
        % Create the finising reports
        createfinishingreports_pdf();
        if values.FinRepLabels == 1
            % Print the labels
            printlabels();
            % Print the finishing reports
%             printfinishingreports_pdf();
        end
    end
end


if (values.SortOutDelNoteWSxml == 1 || values.FDAlabels == 1 || values.ManDelNote == 1) && values.NewCaseIDsPresent == 1 % || values.DailySnapshot == 1
    readcasedata();
    if values.DailySnapshot == 0
        createlabels_fda();
    end
end
% Create sorted out input file for UPS shipments
if values.SortOutDelNoteWSxml == 1
    if values.NewCaseIDsPresent == 1
        % Write them to the UPS Excel
        createshipmentoverview('RegularShipments','0','0','0');
    end
    % Read UPS file again, because there are new shipment lines in it.
    readUPSfile(2);
    % Create input file for UPS WorldShip
    createworldshipinput2();
    % Create input file for delivery note (UPS delivery)
    %createdeliverynoteups();
    createdeliverynote_format2019('UpsShipment');
    % Print the documents
    printshipmentdocuments()
end

% Create input file for delivery note (manual delivery)
if values.ManDelNote == 1
    if values.NewCaseIDsPresent == 1
        createshipmentoverview('RegularShipments','0','0','0');
    end
    % Read UPS file again, because there are new shipment lines in it.
    readUPSfile(2);
    createdeliverynote_format2019('ManualShipment');
end

if values.OpenLogFile == 1
    openlogfile();
end

if values.CheckOMSstatus == 1
    checkomsstatus();
end

if values.CreateStockBoxLabel == 1
    createstockboxlabel();
end

% Only allow access for some persons
if strcmp(values.Username_str,'Lode') || strcmp(values.Username_str,'Natascha')
    
    % Grab the sales order xmls for specific case IDs
    if values.GrabSalesOrdersUser == 1
        grabxmlfilesuser()
    end
    
    % List up all the cases shipped to the production sites for a specific month
    if values.ListShippedCases == 1
        listshippedcases();
    end
    
    % Generate the navision sales shipment xmls for a specific month
    if values.NavSalesShipmentsInput == 1
        createnavisionshipmentinput();
    end
    
%     % Create input files for navision shipments
%     if values.NavShipYear > 1 && values.NavShipMonth > 1 && values.NavShipDayStart > 1
%         % Create input files for navision shipments
%         % readUPSfileUS(2)
%         createnavisionshipmentinput();
%     end
    
    % Create the RS-labels
    if values.GenerateRSLabels == 1
        createlabels_rs()
    end
    
    % Link the invoices to the cases/shipments
    if values.LinkInvoiceToCasesShipments == 1
        linkinvoiceitocasesshipments()
    end
    
        % Link the invoices to the cases/shipments
    if values.GenerateMonthlyStockCorrections == 1
        generatemonthlystockcorrections()
    end
    
end

% Copy all output and temp files to the backup folder.
temp_backup = [values.backupfolder 'MatlabValues\' values.y values.mo values.d '_' values.h values.mi values.s '_' values.Username_str values.TestDebugstr '\Temp'];
mkdir(temp_backup);
copyfile([values.currentversion 'Temp'],temp_backup);
output_backup = [values.backupfolder 'MatlabValues\' values.y values.mo values.d '_' values.h values.mi values.s '_' values.Username_str values.TestDebugstr '\Output'];
mkdir(output_backup);
copyfile([values.currentversion 'Output'],output_backup);

% Log all the actions that have occured: STEP 2 - termination.
if (values.LogEvents == 1 && values.NewCaseIDsPresent == 1) || values.DailySnapshot == 1
    logevents('Termination'); 
    if values.DailySnapshot == 1
        fid2 = fopen([values.currentversion '\Temp\SnapshotDone.txt'],'wt');
        fclose(fid2);
        generateproductionplanning();
    end
end

disp(' ')
disp('Script finished')


% --- Executes on button press in AutoLabels.
function AutoLabels_Callback(hObject, eventdata, handles) %#ok<*DEFNU,*INUSD>
% hObject    handle to AutoLabels (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of AutoLabels

% --- Executes on button press in AutoFinishingReports.
function AutoFinishingReports_Callback(hObject, eventdata, handles)
% hObject    handle to AutoFinishingReports (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of AutoFinishingReports

% --- Executes on button press in ManualLabels.
function ManualLabels_Callback(hObject, eventdata, handles)
% hObject    handle to ManualLabels (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of ManualLabels

% --- Executes on button press in ManualFinishingReports.
function ManualFinishingReports_Callback(hObject, eventdata, handles)
% hObject    handle to ManualFinishingReports (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of ManualFinishingReports


function ManualInsertCaseIDField_Callback(hObject, eventdata, handles)
% hObject    handle to ManualInsertCaseIDField (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of ManualInsertCaseIDField as text
%        str2double(get(hObject,'String')) returns contents of ManualInsertCaseIDField as a double


% --- Executes during object creation, after setting all properties.
function ManualInsertCaseIDField_CreateFcn(hObject, eventdata, handles)
% hObject    handle to ManualInsertCaseIDField (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes during object creation, after setting all properties.
function figure1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to figure1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called


% --- Executes on button press in checkbox17.
function checkbox17_Callback(hObject, eventdata, handles)
% hObject    handle to checkbox17 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of checkbox17


% --- Executes on selection change in listbox2.
function listbox2_Callback(hObject, eventdata, handles)
% hObject    handle to listbox2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns listbox2 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from listbox2


% --- Executes during object creation, after setting all properties.
function listbox2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to listbox2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in MenuYear.
function MenuYear_Callback(hObject, eventdata, handles)
% hObject    handle to MenuYear (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns MenuYear contents as cell array
%        contents{get(hObject,'Value')} returns selected item from MenuYear


% --- Executes during object creation, after setting all properties.
function MenuYear_CreateFcn(hObject, eventdata, handles)
% hObject    handle to MenuYear (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in MenuMonth.
function MenuMonth_Callback(hObject, eventdata, handles)
% hObject    handle to MenuMonth (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns MenuMonth contents as cell array
%        contents{get(hObject,'Value')} returns selected item from MenuMonth


% --- Executes during object creation, after setting all properties.
function MenuMonth_CreateFcn(hObject, eventdata, handles)
% hObject    handle to MenuMonth (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in CreateInvoices.
function CreateInvoices_Callback(hObject, eventdata, handles)
% hObject    handle to CreateInvoices (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of CreateInvoices


% --- Executes on selection change in MenuInvoiceStatus.
function MenuInvoiceStatus_Callback(hObject, eventdata, handles)
% hObject    handle to MenuInvoiceStatus (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns MenuInvoiceStatus contents as cell array
%        contents{get(hObject,'Value')} returns selected item from MenuInvoiceStatus


% --- Executes during object creation, after setting all properties.
function MenuInvoiceStatus_CreateFcn(hObject, eventdata, handles)
% hObject    handle to MenuInvoiceStatus (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit_invoicemanual_Callback(hObject, eventdata, handles)
% hObject    handle to edit_invoicemanual (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_invoicemanual as text
%        str2double(get(hObject,'String')) returns contents of edit_invoicemanual as a double


% --- Executes during object creation, after setting all properties.
function edit_invoicemanual_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_invoicemanual (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in MenuSelectPrinter.
function MenuSelectPrinter_Callback(hObject, eventdata, handles)
% hObject    handle to MenuSelectPrinter (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns MenuSelectPrinter contents as cell array
%        contents{get(hObject,'Value')} returns selected item from MenuSelectPrinter


% --- Executes during object creation, after setting all properties.
function MenuSelectPrinter_CreateFcn(hObject, eventdata, handles)
% hObject    handle to MenuSelectPrinter (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit_InitialInvoiceNumber_Callback(hObject, eventdata, handles)
% hObject    handle to edit_InitialInvoiceNumber (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_InitialInvoiceNumber as text
%        str2double(get(hObject,'String')) returns contents of edit_InitialInvoiceNumber as a double


% --- Executes during object creation, after setting all properties.
function edit_InitialInvoiceNumber_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_InitialInvoiceNumber (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function CaseidsShipments_Callback(hObject, eventdata, handles)
% hObject    handle to CaseidsShipments (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of CaseidsShipments as text
%        str2double(get(hObject,'String')) returns contents of CaseidsShipments as a double


% --- Executes during object creation, after setting all properties.
function CaseidsShipments_CreateFcn(hObject, eventdata, handles)
% hObject    handle to CaseidsShipments (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit_InvoiceYear_Callback(hObject, eventdata, handles)
% hObject    handle to edit_InvoiceYear (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_InvoiceYear as text
%        str2double(get(hObject,'String')) returns contents of edit_InvoiceYear as a double


% --- Executes during object creation, after setting all properties.
function edit_InvoiceYear_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_InvoiceYear (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit_InvoiceMonth_Callback(hObject, eventdata, handles)
% hObject    handle to edit_InvoiceMonth (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_InvoiceMonth as text
%        str2double(get(hObject,'String')) returns contents of edit_InvoiceMonth as a double


% --- Executes during object creation, after setting all properties.
function edit_InvoiceMonth_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_InvoiceMonth (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit_InvoiceDay_Callback(hObject, eventdata, handles)
% hObject    handle to edit_InvoiceDay (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_InvoiceDay as text
%        str2double(get(hObject,'String')) returns contents of edit_InvoiceDay as a double


% --- Executes during object creation, after setting all properties.
function edit_InvoiceDay_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_InvoiceDay (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu_InvoiceYear.
function popupmenu_InvoiceYear_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_InvoiceYear (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_InvoiceYear contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_InvoiceYear


% --- Executes during object creation, after setting all properties.
function popupmenu_InvoiceYear_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_InvoiceYear (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu_InvoiceMonth.
function popupmenu_InvoiceMonth_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_InvoiceMonth (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_InvoiceMonth contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_InvoiceMonth


% --- Executes during object creation, after setting all properties.
function popupmenu_InvoiceMonth_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_InvoiceMonth (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu_InvoiceDay.
function popupmenu_InvoiceDay_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_InvoiceDay (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_InvoiceDay contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_InvoiceDay


% --- Executes during object creation, after setting all properties.
function popupmenu_InvoiceDay_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_InvoiceDay (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function CaseidsDeliveryDocument_Callback(hObject, eventdata, handles)
% hObject    handle to CaseidsDeliveryDocument (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of CaseidsDeliveryDocument as text
%        str2double(get(hObject,'String')) returns contents of CaseidsDeliveryDocument as a double


% --- Executes during object creation, after setting all properties.
function CaseidsDeliveryDocument_CreateFcn(hObject, eventdata, handles)
% hObject    handle to CaseidsDeliveryDocument (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in CreateWorldShipInput.
function CreateWorldShipInput_Callback(hObject, eventdata, handles)
% hObject    handle to CreateWorldShipInput (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of CreateWorldShipInput


% --- Executes on selection change in CaseidsDeliveryDocument_nrcopies.
function CaseidsDeliveryDocument_nrcopies_Callback(hObject, eventdata, handles)
% hObject    handle to CaseidsDeliveryDocument_nrcopies (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns CaseidsDeliveryDocument_nrcopies contents as cell array
%        contents{get(hObject,'Value')} returns selected item from CaseidsDeliveryDocument_nrcopies


% --- Executes during object creation, after setting all properties.
function CaseidsDeliveryDocument_nrcopies_CreateFcn(hObject, eventdata, handles)
% hObject    handle to CaseidsDeliveryDocument_nrcopies (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in CreateUpsDeliveryNote.
function CreateUpsDeliveryNote_Callback(hObject, eventdata, handles)
% hObject    handle to CreateUpsDeliveryNote (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of CreateUpsDeliveryNote



function SpecialShipment_DemoInsoles_Callback(hObject, eventdata, handles)
% hObject    handle to SpecialShipment_DemoInsoles (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of SpecialShipment_DemoInsoles as text
%        str2double(get(hObject,'String')) returns contents of SpecialShipment_DemoInsoles as a double


% --- Executes during object creation, after setting all properties.
function SpecialShipment_DemoInsoles_CreateFcn(hObject, eventdata, handles)
% hObject    handle to SpecialShipment_DemoInsoles (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function SpecialShipment_TopLayers_Callback(hObject, eventdata, handles)
% hObject    handle to SpecialShipment_TopLayers (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of SpecialShipment_TopLayers as text
%        str2double(get(hObject,'String')) returns contents of SpecialShipment_TopLayers as a double


% --- Executes during object creation, after setting all properties.
function SpecialShipment_TopLayers_CreateFcn(hObject, eventdata, handles)
% hObject    handle to SpecialShipment_TopLayers (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function SpecialShipment_Marketing_Callback(hObject, eventdata, handles)
% hObject    handle to SpecialShipment_Marketing (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of SpecialShipment_Marketing as text
%        str2double(get(hObject,'String')) returns contents of SpecialShipment_Marketing as a double


% --- Executes during object creation, after setting all properties.
function SpecialShipment_Marketing_CreateFcn(hObject, eventdata, handles)
% hObject    handle to SpecialShipment_Marketing (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function SpecialShipment_Footscan_Callback(hObject, eventdata, handles)
% hObject    handle to SpecialShipment_Footscan (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of SpecialShipment_Footscan as text
%        str2double(get(hObject,'String')) returns contents of SpecialShipment_Footscan as a double


% --- Executes during object creation, after setting all properties.
function SpecialShipment_Footscan_CreateFcn(hObject, eventdata, handles)
% hObject    handle to SpecialShipment_Footscan (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function SpecialShipment_RegularInsoles_Callback(hObject, eventdata, handles)
% hObject    handle to SpecialShipment_RegularInsoles (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of SpecialShipment_RegularInsoles as text
%        str2double(get(hObject,'String')) returns contents of SpecialShipment_RegularInsoles as a double


% --- Executes during object creation, after setting all properties.
function SpecialShipment_RegularInsoles_CreateFcn(hObject, eventdata, handles)
% hObject    handle to SpecialShipment_RegularInsoles (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in MenuUsername.
function MenuUsername_Callback(hObject, eventdata, handles)
% hObject    handle to MenuUsername (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns MenuUsername contents as cell array
%        contents{get(hObject,'Value')} returns selected item from MenuUsername


% --- Executes during object creation, after setting all properties.
function MenuUsername_CreateFcn(hObject, eventdata, handles)
% hObject    handle to MenuUsername (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in NavSalesOrderInput.
function NavSalesOrderInput_Callback(hObject, eventdata, handles)
% hObject    handle to NavSalesOrderInput (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of NavSalesOrderInput



function NavSalesOrderInput_CaseID_Callback(hObject, eventdata, handles)
% hObject    handle to NavSalesOrderInput_CaseID (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of NavSalesOrderInput_CaseID as text
%        str2double(get(hObject,'String')) returns contents of NavSalesOrderInput_CaseID as a double


% --- Executes during object creation, after setting all properties.
function NavSalesOrderInput_CaseID_CreateFcn(hObject, eventdata, handles)
% hObject    handle to NavSalesOrderInput_CaseID (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in NavShipmentInput.
function NavShipmentInput_Callback(hObject, eventdata, handles)
% hObject    handle to NavShipmentInput (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of NavShipmentInput



function NavisionStartpointShipment_Callback(hObject, eventdata, handles)
% hObject    handle to NavisionStartpointShipment (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of NavisionStartpointShipment as text
%        str2double(get(hObject,'String')) returns contents of NavisionStartpointShipment as a double


% --- Executes during object creation, after setting all properties.
function NavisionStartpointShipment_CreateFcn(hObject, eventdata, handles)
% hObject    handle to NavisionStartpointShipment (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in LimitOnlineReport.
function LimitOnlineReport_Callback(hObject, eventdata, handles)
% hObject    handle to LimitOnlineReport (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of LimitOnlineReport


% --- Executes on selection change in popupmenu_NavShipYear.
function popupmenu_NavShipYear_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_NavShipYear (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_NavShipYear contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_NavShipYear


% --- Executes during object creation, after setting all properties.
function popupmenu_NavShipYear_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_NavShipYear (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu_NavShipMonth.
function popupmenu_NavShipMonth_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_NavShipMonth (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_NavShipMonth contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_NavShipMonth


% --- Executes during object creation, after setting all properties.
function popupmenu_NavShipMonth_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_NavShipMonth (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu_NavShipDayStart.
function popupmenu_NavShipDayStart_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_NavShipDayStart (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_NavShipDayStart contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_NavShipDayStart


% --- Executes during object creation, after setting all properties.
function popupmenu_NavShipDayStart_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_NavShipDayStart (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu_NavShipDayEnd.
function popupmenu_NavShipDayEnd_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_NavShipDayEnd (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_NavShipDayEnd contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_NavShipDayEnd


% --- Executes during object creation, after setting all properties.
function popupmenu_NavShipDayEnd_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_NavShipDayEnd (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu_RSlabelYear.
function popupmenu_RSlabelYear_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_RSlabelYear (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_RSlabelYear contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_RSlabelYear


% --- Executes during object creation, after setting all properties.
function popupmenu_RSlabelYear_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_RSlabelYear (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu_RSlabelMonth.
function popupmenu_RSlabelMonth_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_RSlabelMonth (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_RSlabelMonth contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_RSlabelMonth


% --- Executes during object creation, after setting all properties.
function popupmenu_RSlabelMonth_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_RSlabelMonth (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu_RSlabelBegin.
function popupmenu_RSlabelBegin_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_RSlabelBegin (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_RSlabelBegin contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_RSlabelBegin


% --- Executes during object creation, after setting all properties.
function popupmenu_RSlabelBegin_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_RSlabelBegin (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu_RSlabelFinal.
function popupmenu_RSlabelFinal_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_RSlabelFinal (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_RSlabelFinal contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_RSlabelFinal


% --- Executes during object creation, after setting all properties.
function popupmenu_RSlabelFinal_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_RSlabelFinal (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu_DesiredFunction.
function popupmenu_DesiredFunction_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_DesiredFunction (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_DesiredFunction contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_DesiredFunction


% --- Executes during object creation, after setting all properties.
function popupmenu_DesiredFunction_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_DesiredFunction (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit_NewCaseID_Callback(hObject, eventdata, handles)
% hObject    handle to edit_NewCaseID (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_NewCaseID as text
%        str2double(get(hObject,'String')) returns contents of edit_NewCaseID as a double


% --- Executes during object creation, after setting all properties.
function edit_NewCaseID_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_NewCaseID (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in checkbox_testdebug.
function checkbox_testdebug_Callback(hObject, eventdata, handles)
% hObject    handle to checkbox_testdebug (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of checkbox_testdebug


% --- Executes on button press in checkbox_DoNotGenerateNavisionInput.
function checkbox_DoNotGenerateNavisionInput_Callback(hObject, eventdata, handles)
% hObject    handle to checkbox_DoNotGenerateNavisionInput (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of checkbox_DoNotGenerateNavisionInput


% --- Executes on button press in checkbox_grabsalesorderxmlsups.
function checkbox_grabsalesorderxmlsups_Callback(hObject, eventdata, handles)
% hObject    handle to checkbox_grabsalesorderxmlsups (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of checkbox_grabsalesorderxmlsups


% --- Executes on button press in checkbox_ListShippedCases.
function checkbox_ListShippedCases_Callback(hObject, eventdata, handles)
% hObject    handle to checkbox_ListShippedCases (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of checkbox_ListShippedCases


% --- Executes on button press in checkbox_grabsalesorderxmlsuserinput.
function checkbox_grabsalesorderxmlsuserinput_Callback(hObject, eventdata, handles)
% hObject    handle to checkbox_grabsalesorderxmlsuserinput (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of checkbox_grabsalesorderxmlsuserinput


% --- Executes on selection change in popupmenu_NavisionFunction.
function popupmenu_NavisionFunction_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_NavisionFunction (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_NavisionFunction contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_NavisionFunction


% --- Executes during object creation, after setting all properties.
function popupmenu_NavisionFunction_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_NavisionFunction (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in checkbox_nofinrep.
function checkbox_nofinrep_Callback(hObject, eventdata, handles)
% hObject    handle to checkbox_nofinrep (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of checkbox_nofinrep
