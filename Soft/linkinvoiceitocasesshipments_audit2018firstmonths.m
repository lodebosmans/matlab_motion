
clear all
close all
clc

[y,mo,d,h,mi,s] = getdate();
values.y = y;
values.mo = mo;
values.d = d;
values.h = h;
values.mi = mi;
values.s = s;
values.mainfolder = 'C:\Users\mathlab\Documents\Matlab\'; % Use full path and include \ at the end
values.baseversion = [values.mainfolder '20190318_Interface\']; % Use full path and include \ at the end
values.currentversion = [values.mainfolder '20190318_Interface\']; % Use full path and include \ at the end
values.upsfilename = '20190327_OverviewShipments_v10.xlsx';
values.upsfilepath = 'O:\StandardOperations\UPS\RS Print\'; % Use full path and include \ at the end
values.backupfolder = 'O:\Administration\BackupFiles\'; % Use full path and include \ at the end
values.rsfilesfolder = [values.backupfolder 'RSFiles\']; % Use full path and include \ at the end
values.onlinereportfolder = 'O:\Administration\BackupFiles\OnlineReport\'; % Use full path and include \ at the end
values.onlinereportfilename = '20181206_v49_RSPrint_OR_new.xlsx';
values.onlinereport = [values.onlinereportfolder values.onlinereportfilename]; % Use full path and include \ at the end
values.consultationfolder = 'O:\StandardOperations\OnlineReport\'; % Use full path and include \ at the end
values.salesordercheck_folder = [values.backupfolder 'Navision\SOC\']; % Use full path and include \ at the end
%values.salesordercheck_file = 'salesordercheck_v3'; % Name is used in too many regulare lines to be replaced..
values.salesordercheckuptodate = 'SOC_v3_UpToDate.txt';
values.logfolder = [values.backupfolder 'Logfiles\']; % Use full path and include \ at the end
values.navfolder = [values.backupfolder 'Navision\']; % Use full path and include \ at the end
values.agreementsfile = '20190312_OverviewAgreements.xlsx';
values.shipmentexcelfolder = [values.backupfolder 'ShipmentExcels\']; % Use full path and include \ at the end
values.invoicerequired = 0;
values.rushcasesfolder = 'O:\StandardOperations\Production\'; % Use full path and include \ at the end
values.rushcasesfile = 'RushCases.xlsx';
values.invoicesfolder = [values.navfolder 'Invoices\']; % Use full path and include \ at the end

copyfile([values.salesordercheck_folder 'salesordercheck_v3.mat'],'Temp\salesordercheck_v3.mat'); % always keep full pathway
load Temp\salesordercheck_v3.mat salesordercheck_v3
% Take pre-backup
copyfile([values.salesordercheck_folder 'salesordercheck_v3.mat'],['Output\Navision\' values.y values.mo values.d '_' values.h values.mi values.s '_pre_salesordercheck_v3.mat']);

col_invoicenumber = catchcolumnindex({'InvoiceNumber'},salesordercheck_v3.headers,1);
col_invoicenumber = cell2mat(col_invoicenumber(2,1));
%{'LocationShipment','ShipmentDate','ShipmentService','TrackingNumber','ReferenceRSPrint','ReshipmentInfo','InvoiceNumber'}

disp('Linking invoices to cases - please wait');

% --------------------------------------------------------------------------------

% Make the link between the Excel invoices and the Navision invoices
% Get the file names of the invoices
files = dir([values.invoicesfolder 'New\OLD doc\redundant\*.xlsx']);
nroffiles = length(files);
link = cell(0);
linknr = 0;

if nroffiles == 0
    disp('There are no invoices present to link to the individual cases.');
elseif nroffiles > 0
    % For every invoice
    for k = 1:nroffiles
        % Get filename and date
        info.filename = files(k).name;
        info.year = info.filename(1:4);
        info.month = info.filename(5:6);
        % See if it is already present or not
        [info] = addmoreinfo(info,'redundant');
        if info.alreadypresent == 0
            
            % Read the file and link to cases
            % -------------------------------
            % Get the sheetnames
            [brol,sheets] = xlsfinfo(info.newfile);
            nrofsheets = size(sheets,2);
            % For every sheet
            for sheetnr = 1:nrofsheets
                linknr = linknr + 1;
                % Get the content
                [num,text,raw] = xlsread(info.newfile,sheetnr);
                % Get the invoice number   
                nrofcols = size(text,2);
                goon = 0;
                while goon == 0
                    for row = 3:4
                        for col = 1:nrofcols
                            temp = char(text(row,col));
                            if isempty(temp) == 0 && (strcmp(temp(1:2),'S1') == 1 || strcmp(temp(1:2),'S2'))
                                cinvoice_nr = temp;
                                goon = 1;
                            end
                        end
                    end
                end
                clear goon  
                % Get the external document number
                nrofcols = size(text,2);
                goon = 0;
                while goon == 0
                    for row = 4:20
                        for col = 1:nrofcols
                            temp = char(text(row,col));
                            if isempty(temp) == 0 && strcmp(temp(1:3),'VKF') == 1 
                                VKF_nr = temp;
                                goon = 1;
                            end
                        end
                    end
                end
                clear goon  
                disp(['New link made for ' info.month '/' info.year ' for ' num2str(sheetnr) '/' num2str(nrofsheets) ' :' VKF_nr ' = ' cinvoice_nr])
                link(linknr,1) = {VKF_nr};
                link(linknr,2) = {cinvoice_nr};
                clear VKF_nr
                clear cinvoice_nr
            end
        else
            disp(['Invoice ' info.filename ' is already present in the sorted folder.']);
        end
    end
else
    disp('There seems to be something wrong with the linking of the invoices to the cases..');
end














% -----------------------------------------------------------------------------

% Get the file names of the invoices
files = dir([values.invoicesfolder 'New\OLD doc\*.xlsx']);
nroffiles = length(files);

if nroffiles == 0
    disp('There are no invoices present to link to the individual cases.');
elseif nroffiles > 0
    % For every invoice
    for k = 1:nroffiles
        % Get filename and date
        info.filename = files(k).name;
        info.year = info.filename(1:4);
        info.month = info.filename(5:6);
        % See if it is already present or not
        [info] = addmoreinfo(info,'prenavisionera');
        if info.alreadypresent == 0
            
            % Read the file and link to cases
            % -------------------------------
            
            % Get the content
            [num,text,raw] = xlsread(info.newfile,1);
            % Get the invoice number
            cinvoice_nr = char(text(6,2));
            try
                temp = catchrowindex({cinvoice_nr},link,1);
                temp = cell2mat(temp(2,1));
                temp = char(link(temp,2));
                cinvoice_nr = [cinvoice_nr ' = ' temp];
            catch
            end
            % Get the invoice company name          
            ccus = char(text(4,2));            

            disp(['Linking invoice ' cinvoice_nr '(' num2str(k) '/' num2str(nroffiles) ') of ' info.month '/' info.year ' for ' ccus ' to case IDs - please wait']);
            % Go over the entire invoice to fetch the case IDs
            cinvoice_length = size(text,1);
            for row = 1:cinvoice_length                
                ccid = char(text(row,2));
                if length(ccid) == 12 && (strcmp(ccid(1:3),'RS1') == 1 || strcmp(ccid(1:3),'RS2') == 1)
                    disp([ccid ' - ' cinvoice_nr]);
                    % Find the case ID in the salesordercheck_v3 variable
                    try
                        row_ccid = catchrowindex({ccid},salesordercheck_v3.data,1);
                        row_ccid = cell2mat(row_ccid(2,1));
                        salesordercheck_v3.data(row_ccid,col_invoicenumber) = {cinvoice_nr};
                    catch
                    end
                    
                end
            end
            
            
%             % Move the invoice file
%             % ---------------------
%             % Check if all necessary folders exist
%             if isfolder(info.destinationfolder_sorted) == 0
%                 mkdir(info.destinationfolder_sorted);
%             else
%                 % Folder exists, do nothing.
%             end
%             % Copy to the all folder
%             copyfile(info.newfile,[info.destinationfolder_sorted '\' info.filename])
            
        else
            disp(['Invoice ' info.filename ' is already present in the sorted folder.']);
        end
    end
else
    disp('There seems to be something wrong with the linking of the invoices to the cases..');
end

% save([values.salesordercheck_folder '\salesordercheck_v3.mat'],'salesordercheck_v3')
copyfile([values.salesordercheck_folder '\salesordercheck_v3.mat'],['Output\Navision\' values.y values.mo values.d '_' values.h values.mi values.s '_post_salesordercheck_v3.mat']);
xlswrite([values.salesordercheck_folder '\salesordercheck_v3.xlsx'],salesordercheck_v3.data);
