function linkinvoiceitocasesshipments()

load Temp\values.mat values
copyfile([values.salesordercheck_folder 'salesordercheck_v3.mat'],'Temp\salesordercheck_v3.mat'); % always keep full pathway
load Temp\salesordercheck_v3.mat salesordercheck_v3
% Take pre-backup
copyfile([values.salesordercheck_folder 'salesordercheck_v3.mat'],['Output\Navision\' values.y values.mo values.d '_' values.h values.mi values.s '_pre_salesordercheck_v3.mat']);

col_invoicenumber = catchcolumnindex({'InvoiceNumber'},salesordercheck_v3.headers,1);
col_invoicenumber = cell2mat(col_invoicenumber(2,1));
%{'LocationShipment','ShipmentDate','ShipmentService','TrackingNumber','ReferenceRSPrint','ReshipmentInfo','InvoiceNumber'}

disp('Linking invoices to cases - please wait');

% Get the file names of the invoices
files = dir([values.invoicesfolder 'New\*.xlsx']);
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
        [info] = addmoreinfo(info,'linkinvoiceitocasesshipments');
        if info.alreadypresent == 0
            
            % Read the file and link to cases
            % -------------------------------
            % Get the sheetnames
            [brol,sheets] = xlsfinfo(info.newfile);
            nrofsheets = size(sheets,2);
            % For every sheet
            for sheetnr = 1:nrofsheets
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
                % Get the invoice company name
                nrofcols = size(text,2);
                goon = 0;
                while goon == 0
                    for col = 5:nrofcols
                        if isempty(char(text(23,col))) == 0
                            ccus = char(text(23,col));
                            goon = 1;
                        end
                    end
                end
                clear goon                
                disp(['Linking invoice ' cinvoice_nr ' (' num2str(sheetnr) '/' num2str(nrofsheets) ') of ' info.month '/' info.year ' for ' ccus ' to case IDs - please wait']);                
                % Go over the entire invoice to fetch the case IDs
                cinvoice_length = size(text,1);
                cinvoice_width = size(text,2);
                for row = 1:cinvoice_length
                    for col = 3:cinvoice_width
                        ccid = char(text(row,col));
                        if length(ccid) == 12 && (strcmp(ccid(1:3),'RS1') == 1 || strcmp(ccid(1:3),'RS2') == 1)
                            disp([ccid ' - ' cinvoice_nr]);
                            % Find the case ID in the salesordercheck_v3 variable
                            row_ccid = catchrowindex({ccid},salesordercheck_v3.data,1);
                            row_ccid = cell2mat(row_ccid(2,1));
                            % Inject the invoice number
                            salesordercheck_v3.data(row_ccid,col_invoicenumber) = {cinvoice_nr};
                        end
                    end                    
                end
            end
            
            % Move the invoice file
            % ---------------------
            % Check if all necessary folders exist
            if isfolder(info.destinationfolder_sorted) == 0
                mkdir(info.destinationfolder_sorted);
            else
                % Folder exists, do nothing.
            end
            % Copy to the all folder
            movefile(info.newfile,[info.destinationfolder_sorted '\' info.filename])
            
        else
            disp(['Invoice ' info.filename ' is already present in the sorted folder.']);
        end
    end
else
    disp('There seems to be something wrong with the linking of the invoices to the cases..');
end

save([values.salesordercheck_folder '\salesordercheck_v3.mat'],'salesordercheck_v3')
copyfile([values.salesordercheck_folder '\salesordercheck_v3.mat'],['Output\Navision\' values.y values.mo values.d '_' values.h values.mi values.s '_post_salesordercheck_v3.mat']);
xlswrite([values.salesordercheck_folder '\salesordercheck_v3.xlsx'],salesordercheck_v3.data);

end