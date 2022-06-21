function createworldshipinput2()

load Temp\values.mat values
load Temp\UPSfile_shipment.mat UPSfile_shipment
load Temp\UPSfile_product.mat UPSfile_product %#ok<NASGU>
load Temp\UPSfile_itemnumbers.mat UPSfile_itemnumbers
load Temp\customers.mat customers
load([values.salesordercheck_folder 'salesordercheck_v3.mat'],'salesordercheck_v3')

disp('Creating WorldShip input - please wait');

[nrofrows,nrofcols] = size(UPSfile_shipment); %#ok<ASGLU>

% col.shipmentlabel = UPSfile_shipment(1,:);
% col.itemnrlabel = UPSfile_itemnumbers(1,:);

%shipmentcounter = 0;
% Shipment headers
col_shipped = catchcolumnindex({'PromotedOMS'},UPSfile_shipment,1);
col_shipped = cell2mat(col_shipped(2,1));
col_service = catchcolumnindex({'Service'},UPSfile_shipment,1);
col_service = cell2mat(col_service(2,1));
col_shipnr = catchcolumnindex({'ShipmentNumber'},UPSfile_shipment,1);
col_shipnr = cell2mat(col_shipnr(2,1));
col_customernumber = catchcolumnindex({'CustomerNumber'},UPSfile_shipment,1);
col_customernumber = cell2mat(col_customernumber(2,1));
col_serialnumbers = catchcolumnindex({'SerialNumbers'},UPSfile_shipment,1);
col_serialnumbers = cell2mat(col_serialnumbers(2,1));
col_nonserialnumbers = catchcolumnindex({'NonSerialNumbers'},UPSfile_shipment,1);
col_nonserialnumbers = cell2mat(col_nonserialnumbers(2,1));
col_ShippedFrom = catchcolumnindex({'ShippedFrom'},UPSfile_shipment,1);
col_ShippedFrom = cell2mat(col_ShippedFrom(2,1));
% Itemnumber headers
col_HarmCode = catchcolumnindex({'Harmonized code'},UPSfile_itemnumbers,1);
col_HarmCode = cell2mat(col_HarmCode(2,1));
col_ItemNumber = catchcolumnindex({'ItemNumber'},UPSfile_itemnumbers,1);
col_ItemNumber = cell2mat(col_ItemNumber(2,1));
col_Description = catchcolumnindex({'Description'},UPSfile_itemnumbers,1);
col_Description = cell2mat(col_Description(2,1));
col_InPrice = catchcolumnindex({'Price/piece (in euro)'},UPSfile_itemnumbers,1);
col_InPrice = cell2mat(col_InPrice(2,1));
% Salesordercheck headers
col_caseID = catchcolumnindex({'CaseID'},salesordercheck_v3.headers,1);
col_caseID = cell2mat(col_caseID(2,1));
col_Price = catchcolumnindex({'Price'},salesordercheck_v3.headers,1);
col_Price = cell2mat(col_Price(2,1));

list_shipments = {};
itemnrs_footscanplates = ["20010000","20020000","20030000","20040000","20050000","20060000","20090000","20110000","20120000", ...
    "20140000","20160000","20190000","20200000","20220000","20230000","20240000","20250000","20260000","20270000","20280000", ...
    "20290000","20300000","20310000","20320000","20330000","20350000","20360000","20370000"];

itemnrs_psl = ["50000001","50000002","50000003","50000004","50000005","50000010","50000013","50000025","50000042","50000047", ...
    "50000090","50000103","50000104","50000109","50000111"];

for cr = 2:nrofrows
    %     if cr == 2848
    %         stopit = 1;
    %     end
    if strcmp(UPSfile_shipment(cr,col_shipnr),'-') == 1
        hasgotshipnr = 0;
    else
        hasgotshipnr = 1;
    end
    % Get the customernumber first
    cstn_c = char(UPSfile_shipment(cr,col_customernumber));
    cstn_c_ = strrep(cstn_c,'-','_');
    
    importerpresent = 0;
        
    % Check is the shipment in the row is already shipped. If not, fetch all necessary data
    if strcmp(cell2mat(UPSfile_shipment(cr,col_shipped)),'-') == 1 && strcmp(UPSfile_shipment(cr,col_service),'UPS') == 1 ...
            && hasgotshipnr == 1 && strcmp(char(UPSfile_shipment(cr,col_ShippedFrom)),values.CountryOfShipment) == 1
        
        % Fetch the data
        % --------------
        %shipmentcounter = shipmentcounter + 1;
        shipmentnumber = num2str(cell2mat((UPSfile_shipment(cr,col_shipnr))));
      
        % Add to the list
        nrofshipments = size(list_shipments,1);
        list_shipments(nrofshipments+1,1) =   {[values.y values.mo values.d '_' values.h values.mi values.s  '_Shipment_' shipmentnumber '_' customers.(cstn_c_).Company]};
        
        % ShipTo
        % ------
        %temp = strrep(temp,'&','and');
        Shipment.ShipTo.CompanyOrName = customers.(cstn_c_).Company;
        Shipment.ShipTo.CustomerNumber = cstn_c;
        Shipment.ShipTo.CustomerNumber_ = cstn_c_;
        disp(['Currently processing row ' num2str(cr) ' for shipment ' num2str(cell2mat(UPSfile_shipment(cr,4))) ' to ' char(Shipment.ShipTo.CompanyOrName) ' (' char(UPSfile_shipment(cr,6)) ')' ]);
        Shipment.ShipTo.Attention = customers.(cstn_c_).DeliveryAddress.DeliveryContactPerson;
        Shipment.ShipTo.Address1 = customers.(cstn_c_).DeliveryAddress.DeliveryAddress1;
        Shipment.ShipTo.Address2 = customers.(cstn_c_).DeliveryAddress.DeliveryAddress2;
        Shipment.ShipTo.Address3 = customers.(cstn_c_).DeliveryAddress.DeliveryAddress3;
        Shipment.ShipTo.CountryTerritory = customers.(cstn_c_).DeliveryAddress.DeliveryCountryCode;
        Shipment.ShipTo.PostalCode = customers.(cstn_c_).DeliveryAddress.DeliveryPostalCode;
        Shipment.ShipTo.CityOrTown = customers.(cstn_c_).DeliveryAddress.DeliveryCity;
        Shipment.ShipTo.StateProvinceCounty = customers.(cstn_c_).DeliveryAddress.DeliveryState;
        Shipment.ShipTo.StateProvinceCounty2 = customers.(cstn_c_).DeliveryAddress.DeliveryStateCode;
        Shipment.ShipTo.Telephone = customers.(cstn_c_).DeliveryAddress.DeliveryPhone;
        Shipment.ShipTo.FaxNumber = '';
        Shipment.ShipTo.EmailAddress = customers.(cstn_c_).Email;
        Shipment.ShipTo.NotificationEmailAddress = customers.(cstn_c_).NotificationEmail;
        Shipment.ShipTo.TaxIDNumber = customers.(cstn_c_).VatEin;
        
        % Importer
        % --------
        % See nonEU section for the rest
        
        % ShipmentInformation
        % -------------------
        temp = catchcolumnindex({'ServiceCode'},UPSfile_shipment,1);
        Shipment.ShipmentInformation.ServiceType = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        Shipment.ShipmentInformation.PackageType = 'CP';
        temp = catchcolumnindex({'NrPackages'},UPSfile_shipment,1);
        Shipment.ShipmentInformation.NumberOfPackages = num2str(cell2mat(UPSfile_shipment(cr,cell2mat(temp(2,1)))));
        temp = catchcolumnindex({'Weight'},UPSfile_shipment,1);
        Shipment.ShipmentInformation.ShipmentActualWeight = num2str(cell2mat(UPSfile_shipment(cr,cell2mat(temp(2,1)))));
        temp = catchcolumnindex({'ContentDescription'},UPSfile_shipment,1);
        Shipment.ShipmentInformation.DescriptionOfGoods = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = catchcolumnindex({'Reference'},UPSfile_shipment,1);
        Shipment.ShipmentInformation.Reference1 = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        Shipment.ShipmentInformation.Reference2 = '';
        Shipment.ShipmentInformation.ShipperNumber = '';
        Shipment.ShipmentInformation.BillingOption = 'PP';
        Shipment.ShipmentInformation.ProcessAsPaperless = 'Y';
        temp = catchcolumnindex({'BillTransportationTo'},UPSfile_shipment,1);
        Shipment.ShipmentInformation.BillTransportationTo = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = catchcolumnindex({'DutyTax'},UPSfile_shipment,1);
        Shipment.ShipmentInformation.BillDutyTaxTo = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        % If there is an importer, overrule the info from the UPS excel
        if isfield(customers.(cstn_c_), 'Importer')
            if strcmp(customers.(cstn_c_).Importer,'-') == 0
                % do something
                Shipment.ShipmentInformation.BillDutyTaxTo = 'TP';
            end
        end
        
        % QVNOption
        % ---------
        
        % QVNRecipientAndNotificationTypes
        % --------------------------------
        Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.ContactName = Shipment.ShipTo.Attention;
        Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.NotificationEMailAddress = Shipment.ShipTo.NotificationEmailAddress;
        Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.Ship = '1';
        Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.Exception = '0';
        Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.Delivery = '1';
        
        % QVNOption (part 2)
        % ------------------
        if strcmp(values.CountryOfShipment,'BE') == 1
            Shipment.ShipmentInformation.QVNOption.ShipFromCompanyOrName = 'RS Print (Paal-BE)';
            currency = 'EUR';
        elseif strcmp(values.CountryOfShipment,'US') == 1
            Shipment.ShipmentInformation.QVNOption.ShipFromCompanyOrName = 'RS PRint (Materialise US)';
            currency = 'USD';
        else
            problem = 1;
            clear all
        end
        Shipment.ShipmentInformation.QVNOption.FailedEMailAddress = 'support@rsprint.be';
        Shipment.ShipmentInformation.QVNOption.SubjectLine = ['RS Print shipment: ' Shipment.ShipmentInformation.Reference1];
        temp = catchcolumnindex({'Memo'},UPSfile_shipment,1);
        temp = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = strrep(temp,'   ',''); % Reduce the length of the string
        temp = strrep(temp,'(1)',''); % Reduce the length of the string
        temp = strrep(temp,'-',''); % Idem ditto
        temp = strrep(temp,',,',','); % Idem ditto
        temp = strrep(temp,',,',','); % Idem ditto
        Shipment.ShipmentInformation.QVNOption.Memo = temp;
        
        eu_col = catchcolumnindex({'EU'},UPSfile_shipment,1);
        eu_col = cell2mat(eu_col(2,1));
        if strcmp(char(UPSfile_shipment(cr,eu_col)),'Noooo') == 1 || strcmp(values.CountryOfShipment,'BE') == 1
            
            % Importer
            % --------
            clear test_importer
            try
                test_importer = customers.(cstn_c_).Importer;
            catch
            end
            if exist('test_importer','var') == 1
                if strcmp(customers.(cstn_c_).Importer,'-') == 0
                    importer = customers.(cstn_c_).Importer;
                    importer = strrep(importer,'-','_');
                    importerpresent = 1;
                    
                    Shipment.Importer.CompanyOrName = customers.(importer).Company;
                    Shipment.Importer.Attention = customers.(importer).ContactPerson;
                    Shipment.Importer.Address1 = customers.(importer).InvoiceAddress.InvoiceAddress1;
                    Shipment.Importer.Address2 = customers.(importer).InvoiceAddress.InvoiceAddress2;
                    Shipment.Importer.Address3 = customers.(importer).InvoiceAddress.InvoiceAddress3;
                    Shipment.Importer.CityOrTown = customers.(importer).InvoiceAddress.InvoiceCity;
                    Shipment.Importer.CountryTerritory = customers.(importer).DeliveryAddress.DeliveryCountryCode;
                    Shipment.Importer.PostalCode = customers.(importer).InvoiceAddress.InvoicePostalCode;
                    Shipment.Importer.StateProvinceCounty = customers.(importer).InvoiceAddress.InvoiceState;
                    Shipment.Importer.Telephone = customers.(importer).InvoiceAddress.InvoicePhone;
                    Shipment.Importer.UpsAccountNumber = customers.(importer).UPSAccountNumber;
                    Shipment.Importer.TaxIDNumber = customers.(importer).VatEin;
                end
            end
            
            % International documentation
            % ---------------------------
            Shipment.InternationalDocumentation.InvoiceImporterSameAsShipTo = 'N';
            Shipment.InternationalDocumentation.InvoiceTermOfSale = '';
            Shipment.InternationalDocumentation.InvoiceReasonForExport = '02';
            Shipment.InternationalDocumentation.InvoiceAdditionalComments = '';
            if strcmp(cstn_c(1:5),'C1015') == 1
                Shipment.InternationalDocumentation.InvoiceDeclarationStatement = 'Korea';
                Shipment.InternationalDocumentation.InvoiceAdditionalComments = 'Exporter code (EORI) BE74 0551855071. The exporter of the products covered by this document, declares that except where otherwise clearly indicated, these products are of EU preferential origin. Tom Peeters, RS Print.';
            elseif strcmp(Shipment.ShipTo.CountryTerritory,'GB') == 1
                Shipment.InternationalDocumentation.InvoiceDeclarationStatement = 'UK';
                Shipment.InternationalDocumentation.InvoiceAdditionalComments = '(Exporter Reference No. BEREXBE0551855071) I hereby certify that the information on this invoice is true and correct and the contents and value of this shipment is as stated above.';
            elseif strcmp(Shipment.ShipTo.CountryTerritory,'US') == 4
                Shipment.InternationalDocumentation.InvoiceDeclarationStatement = 'US';
                Shipment.InternationalDocumentation.InvoiceAdditionalComments = 'Materialise FDA registration #: 3003998208';
            else
                Shipment.InternationalDocumentation.InvoiceDeclarationStatement = 'Invoice';
                Shipment.InternationalDocumentation.InvoiceAdditionalComments = 'Exporter code (EORI) BE74 0551855071. The exporter of the products covered by this document, declares that except where otherwise clearly indicated, these products are of EU preferential origin. Tom Peeters, RS Print.';
            end
            Shipment.InternationalDocumentation.InvoiceCurrencyCode = 'EUR';
            %temp = catchcolumnindex({'InvoiceTotal'},col.shipmentlabel,1);
            %Shipment.InternationalDocumentation.InvoiceLineTotal = num2str(cell2mat(UPSfile_shipment(cr,cell2mat(temp(2,1)))));
            
            
            % ----------------------------------------------------------------------------------------------------
            % ----------------------------------------------------------------------------------------------------
            % ----------------------------------------------------------------------------------------------------
            % ----------------------------------------------------------------------------------------------------
            
            
            
            % Goods
            % -----
            content = char(UPSfile_shipment(cr,col_serialnumbers));
            goodscounter = 0;
            InvoiceLineTotal = 0;
            
            if strcmp(content,'-') == 0
                
                % Go over all insoles
                % -------------------
                [content_insoles] = readmanualcaseids({},'none',0,content,1);
                nrofinsoles = size(content_insoles,1);
                for ci = 1:nrofinsoles
                    ccid = content_insoles(ci,col_caseID); % in cell format.
                    goodscounter = goodscounter + 1;
                    goodscounter2 = num2str(goodscounter);
                    if strcmp(Shipment.ShipTo.CountryTerritory,'US') == 1
                        eval(['Shipment.Goods' goodscounter2 '.DescriptionOfGood = ''RS Print powered by Materialise - Orthosis, Corrective Shoe – FDA Product Code: KNP'';']);
                    else
                        eval(['Shipment.Goods' goodscounter2 '.DescriptionOfGood = ''Insoles'';']);
                    end
                    eval(['Shipment.Goods' goodscounter2 '.Inv_NAFTA_TariffCode = num2str(cell2mat((UPSfile_product(3,1))));']);
                    eval(['Shipment.Goods' goodscounter2 '.Inv_NAFTA_CO_CountryTerritoryOfOrigin = values.CountryOfShipment;']);
                    eval(['Shipment.Goods' goodscounter2 '.InvoiceUnits = ''1'';']);
                    eval(['Shipment.Goods' goodscounter2 '.InvoiceUnitOfMeasure = ''PC'';']);
                    % Inject price here. First find it in the salesordercheck_v3.
                    soc_row = catchrowindex(ccid,salesordercheck_v3.data,col_caseID);
                    soc_row = cell2mat(soc_row(2,1));
                    cprice = cell2mat(salesordercheck_v3.data(soc_row,col_Price));
                    InvoiceLineTotal = InvoiceLineTotal + cprice;
                    eval(['Shipment.Goods' goodscounter2 '.Invoice_SED_UnitPrice = num2str(cprice) ;']);
                    eval(['Shipment.Goods' goodscounter2 '.InvoiceCurrencyCode = currency;']);
                end
                
                % Go over all RS-numbers
                % ----------------------
                [content_rsfiles] = readmanualcaseids({},'none',0,content,2);
                nrofrsfiles = size(content_rsfiles,1);
                for crs = 1:nrofrsfiles
                    goodscounter = goodscounter + 1;
                    goodscounter2 = num2str(goodscounter);
                    rs_line.retour = 0;
                    rs_line.footscan = 0;
                    rs_line.scan3D = 0;
                    rs_line.interfacebox = 0;
                    rs_line.flightcase = 0;
                    rs_line.runways = 0;
                    rs_line.dongle = 0;
                    rs_line.marketingpack = 0;
                    rs_line.rma = 0;
                    rs_line.psl = 0;
                    rs_line.other = 0;
                    footscandeclaration = 0;
                    % Get the RS-number
                    ccid = char(content_rsfiles(crs,col_caseID)); % in cell format.
                    
                    % Open the file and determine the content
                    % First tab with the general information
                    % --------------------------------------
                    [brol1,brol2,crsgeninfo] = xlsread([values.rsfilesfolder ccid '.xlsx'],1);
                    row_yourreference = catchrowindex({'Your Reference'},crsgeninfo,1);
                    row_yourreference = cell2mat(row_yourreference(2,1));
                    yourreference = char(crsgeninfo(row_yourreference,2));
                    if contains(upper(yourreference),'RMA')
                       rs_line.rma = 1; 
                    end
                    
                    % Second tab with the content => define the description
                    % -----------------------------------------------------
                    [brol1,brol2,crscontent] = xlsread([values.rsfilesfolder ccid '.xlsx'],2);
                    %crscontent(cellfun(@(x) isnumeric(x) && isnan(x), crscontent)) = {'-'};
                    nrofrowrs = size(crscontent,1);
                    col_ItemNumberRSFile = catchcolumnindex({'No.'},crscontent,2);
                    col_ItemNumberRSFile = cell2mat(col_ItemNumberRSFile(2,1));
                    
                    for crowrs = 3:nrofrowrs
                        temp = char(crscontent(crowrs,col_ItemNumberRSFile));
                        if length(temp) == 0
                            % This is an empty cell. Move on.
                        elseif strcmp(temp,'70000010P') == 1 || strcmp(temp,'70000010R') == 1 || strcmp(temp,'70000010R ') == 1
                            % This is a retour case.
                            rs_line.retour = 1;
                        elseif contains(temp,itemnrs_footscanplates) == 1 || rs_line.rma == 1
                            % This is a footscan plate.
                            rs_line.footscan = 1;
                        elseif strcmp(temp,'91000001') == 1 ...
                                || strcmp(temp,'91000002') == 1 ...
                                || strcmp(temp,'91000004') == 1 ...
                                || strcmp(temp,'91000006') == 1
                            % This is a 3D scanner.
                            rs_line.scan3D = 1;
                        elseif strcmp(temp,'21000100') == 1 ...
                                || strcmp(temp,'21000200') == 1
                            % This is an interface box.
                            rs_line.interfacebox = 1;
                        elseif contains(temp,itemnrs_psl) == 1 
                            % This is a PSL kit.
                            rs_line.psl = 1;
                        elseif strcmp(temp,'30000002') == 1 ...
                                || strcmp(temp,'30000009') == 1 ...
                                || strcmp(temp,'30000025') == 1 ...
                                || strcmp(temp,'30000026') == 1 ...
                                || strcmp(temp,'30000027') == 1
                            % This is a flight case.
                            rs_line.flightcase = 1;
                        elseif strcmp(temp,'60000033') == 1 ...
                                || strcmp(temp,'60000034') == 1 ...
                                || strcmp(temp,'60000035') == 1 ...
                                || strcmp(temp,'60000036') == 1 ...
                                || strcmp(temp,'60000037') == 1 ...
                                || strcmp(temp,'60000038') == 1
                            % These are runways.
                            rs_line.runways = 1;
                        elseif strcmp(temp,'60000047') == 1
                            % This is a dongle.
                            rs_line.dongle = 1;
                        elseif strcmp(temp,'M75000001') == 1 ...
                                || strcmp(temp,'M75000002') == 1 ...
                                || strcmp(temp,'M75000003') == 1 ...
                                || strcmp(temp,'M75999999') == 1
                            % This is a marketing pack.
                            rs_line.marketingpack = 1;
                        else
                            rs_line.other = 1;
                        end
                    end
                    % Now pick the content with the most summarizing value.
                    if rs_line.footscan == 1
                        DescriptionOfGood = 'Footscan pressure plate';
                        Harmcode = '90318038';
                        footscandeclaration = 1;
                    elseif rs_line.scan3D == 1
                        DescriptionOfGood = '3D scanner';
                        Harmcode = '90314990';
                        footscandeclaration = 1;
                    elseif rs_line.interfacebox == 1
                        DescriptionOfGood = 'Interfacebox';
                        Harmcode = '90319085';
                    elseif rs_line.dongle == 1
                        DescriptionOfGood = 'Security dongle';
                        Harmcode = '90319085';
                    elseif rs_line.psl == 1
                        DescriptionOfGood = 'PSL';
                        Harmcode = '90319085';
                    elseif rs_line.flightcase == 1
                        DescriptionOfGood = 'Flightcase';
                        Harmcode = '90319085';
                    elseif rs_line.runways == 1
                        DescriptionOfGood = 'Runways';
                        Harmcode = '90319085';
                    elseif rs_line.marketingpack == 1
                        DescriptionOfGood = 'Marketing pack';
                        Harmcode = '90319085';
                    elseif rs_line.retour == 1
                        DescriptionOfGood = 'Repaired insoles';
                        Harmcode = '9021100090';
                    elseif rs_line.other == 1
                        prompt = {['Enter a description of the content of ' ccid ':'],'Define the Harmonized code:'};
                        title = 'Provide input before proceding';
                        dims = [1 70];
                        definput = {'Top layers','9021100090'};
                        answer = inputdlg(prompt,title,dims,definput);
                        DescriptionOfGood = char(answer(1,1));
                        Harmcode = char(answer(2,1));
                    else
                        % If the code fails here, there is something terribly wrong :)
                        clear all
                    end
                    clear rs_line
                    % Third tab with the price
                    % ------------------------
                    [brol1,brol2,crsvalue] = xlsread([values.rsfilesfolder ccid '.xlsx'],3);
                    cprice = cell2mat(crsvalue(7,2));
                    if cprice == 0
                        cprice = 0.01;
                    end
                    % Don't forget to update the price!
                    InvoiceLineTotal = InvoiceLineTotal + cprice;
                    % Save to shipment variable.
                    eval(['Shipment.Goods' goodscounter2 '.DescriptionOfGood = DescriptionOfGood;']);
                    eval(['Shipment.Goods' goodscounter2 '.Inv_NAFTA_TariffCode = Harmcode;']);
                    eval(['Shipment.Goods' goodscounter2 '.Inv_NAFTA_CO_CountryTerritoryOfOrigin = values.CountryOfShipment;']);
                    eval(['Shipment.Goods' goodscounter2 '.InvoiceUnits = ''1'';']); % Quantity will always be one, as Natascha promised not to include more than one footscan plate per sales order.
                    eval(['Shipment.Goods' goodscounter2 '.InvoiceUnitOfMeasure = ''PC'';']);
                    eval(['Shipment.Goods' goodscounter2 '.Invoice_SED_UnitPrice = num2str(cprice);']);
                    eval(['Shipment.Goods' goodscounter2 '.InvoiceCurrencyCode = currency;']);
                    % Print the declaration file if necessary
                    if footscandeclaration == 1
                        printfootscandeclaration();
                        footscandeclaration = 0;
                    end
                end
            else
                
                % Create the goods for a shipment without case IDs or RS-numbers
                content = char(UPSfile_shipment(cr,col_nonserialnumbers));
                % Replace the non-comma characters
                content = strrep(content,';',',');
                content = strrep(content,'/',',');
                content = strrep(content,',,,',',');
                content = strrep(content,',,',',');
                content = strrep(content,' ','');
                % Count the number of comma's in the string to distinguish the different itemnumbers
                comma = strfind(content, ',');
                openbracket = strfind(content, '(');
                closebracket = strfind(content, ')');
                nrofitems = size(openbracket,2);
                goodscounter = nrofitems;
                for ci = 1:nrofitems
                    % Get the item number and the amount
                    if ci == 1
                        itemnumber = content(1:openbracket(1,ci)-1);
                    elseif ci > 1
                        itemnumber = content(comma(1,ci-1)+1:openbracket(1,ci)-1);
                    end
                    camount = content(openbracket(1,ci)+1:closebracket(1,ci)-1);
                    goodscounter2 = num2str(ci);
                    % Search for the itemprice or ask for it when not available.
                    try
                        % If the itemnumber is found
                        crow = catchrowindex({itemnumber},UPSfile_itemnumbers,col_ItemNumber);
                        crow = cell2mat(crow(2,1));
                        cdescription = char(UPSfile_itemnumbers(crow,col_Description));
                        cunitprice = num2str(cell2mat(UPSfile_itemnumbers(crow,col_InPrice)));
                        charmcode = num2str(cell2mat(UPSfile_itemnumbers(crow,col_HarmCode)));
                        if strcmp(cdescription,'-') == 1 || strcmp(cunitprice,'-') == 1  || strcmp(charmcode,'-') == 1
                            [answer] = getitemdetails('noTLharm',Shipment,itemnumber,camount,cdescription,cunitprice,charmcode);
                            cdescription = char(answer(5,1));
                            cunitprice = char(answer(3,1));
                            charmcode = char(answer(4,1));
                        end
                    catch
                        % If the itemnumber is not present, come ask for the information
                        [answer] = getitemdetails('TLharm',Shipment,itemnumber,camount,'','','9021100090');
                        cdescription = char(answer(5,1));
                        cunitprice = char(answer(3,1));
                        charmcode = char(answer(4,1));
                    end
                    % Write it to the shipment variable.
                    eval(['Shipment.Goods' goodscounter2 '.DescriptionOfGood = cdescription;']);
                    eval(['Shipment.Goods' goodscounter2 '.Inv_NAFTA_TariffCode = charmcode;']);
                    eval(['Shipment.Goods' goodscounter2 '.Inv_NAFTA_CO_CountryTerritoryOfOrigin = values.CountryOfShipment;']);
                    eval(['Shipment.Goods' goodscounter2 '.InvoiceUnits = camount;']);
                    eval(['Shipment.Goods' goodscounter2 '.InvoiceUnitOfMeasure = ''PC'';']);
                    % Inject price here. First find it in the itemnumberlist
                    cprice = str2double(camount) * str2double(cunitprice);
                    InvoiceLineTotal = InvoiceLineTotal + cprice;
                    eval(['Shipment.Goods' goodscounter2 '.Invoice_SED_UnitPrice = cunitprice;']);
                    eval(['Shipment.Goods' goodscounter2 '.InvoiceCurrencyCode = currency;']);
                end
            end
            InvoiceLineTotal = num2str(InvoiceLineTotal);
        end
        % Update the final shipment value
        Shipment.InternationalDocumentation.InvoiceLineTotal = InvoiceLineTotal;
        
        % ----------------------------------------------------------------------------------------------------
        % ----------------------------------------------------------------------------------------------------
        % ----------------------------------------------------------------------------------------------------
        % ----------------------------------------------------------------------------------------------------
        
        
        % Write to file
        
        docNode = com.mathworks.xml.XMLUtils.createDocument('OpenShipments');
        docRootNode = docNode.getDocumentElement;
        docRootNode.setAttribute('xmlns','x-schema:OpenShipments.xdr');
        
        OpenShipment_node = docNode.createElement('OpenShipment');
        %OpenShipment_node_text = docNode.createTextNode('OpenShipment');
        %OpenShipment__node.appendChild(OpenShipment_node_text);
        docRootNode.appendChild(OpenShipment_node);
        OpenShipment_node.setAttribute('ShipmentOption','SC');
        OpenShipment_node.setAttribute('ProcessStatus','');
        
        % ------------------------------------------------------------- %
        %                          ShipTo                               %
        % ------------------------------------------------------------- %
        
        ShipTo_tag = docNode.createElement('ShipTo');
        % ShipTo_text = docNode.createTextNode('ShipTo_text');
        % ShipTo_tag.appendChild(ShipTo_text);
        OpenShipment_node.appendChild(ShipTo_tag);
        
        CompanyOrName_tag = docNode.createElement('CompanyOrName');
        CompanyOrName_text = docNode.createTextNode(char(Shipment.ShipTo.CompanyOrName));
        CompanyOrName_tag.appendChild(CompanyOrName_text);
        ShipTo_tag.appendChild(CompanyOrName_tag);
        
        Attention_tag = docNode.createElement('Attention');
        Attention_text = docNode.createTextNode(char(Shipment.ShipTo.Attention));
        Attention_tag.appendChild(Attention_text);
        ShipTo_tag.appendChild(Attention_tag);
        
        Address1_tag = docNode.createElement('Address1');
        Address1_text = docNode.createTextNode(char(Shipment.ShipTo.Address1));
        Address1_tag.appendChild(Address1_text);
        ShipTo_tag.appendChild(Address1_tag);
        
        Address2_tag = docNode.createElement('Address2');
        Address2_text = docNode.createTextNode(char(Shipment.ShipTo.Address2));
        Address2_tag.appendChild(Address2_text);
        ShipTo_tag.appendChild(Address2_tag);
        
        Address3_tag = docNode.createElement('Address3');
        Address3_text = docNode.createTextNode(char(Shipment.ShipTo.Address3));
        Address3_tag.appendChild(Address3_text);
        ShipTo_tag.appendChild(Address3_tag);
        
        CountryTerritory_tag = docNode.createElement('CountryTerritory');
        CountryTerritory_text = docNode.createTextNode(char(Shipment.ShipTo.CountryTerritory));
        CountryTerritory_tag.appendChild(CountryTerritory_text);
        ShipTo_tag.appendChild(CountryTerritory_tag);
        
        PostalCode_tag = docNode.createElement('PostalCode');
        PostalCode_text = docNode.createTextNode(char(Shipment.ShipTo.PostalCode));
        PostalCode_tag.appendChild(PostalCode_text);
        ShipTo_tag.appendChild(PostalCode_tag);
        
        CityOrTown_tag = docNode.createElement('CityOrTown');
        CityOrTown_text = docNode.createTextNode(char(Shipment.ShipTo.CityOrTown));
        CityOrTown_tag.appendChild(CityOrTown_text);
        ShipTo_tag.appendChild(CityOrTown_tag);
        
        StateProvinceCounty_tag = docNode.createElement('StateProvinceCounty');
        StateProvinceCounty_text = docNode.createTextNode(char(Shipment.ShipTo.StateProvinceCounty2));
        StateProvinceCounty_tag.appendChild(StateProvinceCounty_text);
        ShipTo_tag.appendChild(StateProvinceCounty_tag);
        
        Telephone_tag = docNode.createElement('Telephone');
        Telephone_text = docNode.createTextNode(char(Shipment.ShipTo.Telephone));
        Telephone_tag.appendChild(Telephone_text);
        ShipTo_tag.appendChild(Telephone_tag);
        
        FaxNumber_tag = docNode.createElement('FaxNumber');
        FaxNumber_text = docNode.createTextNode('');
        FaxNumber_tag.appendChild(FaxNumber_text);
        ShipTo_tag.appendChild(FaxNumber_tag);
        
        EmailAddress_tag = docNode.createElement('EmailAddress');
        EmailAddress_text = docNode.createTextNode(char(Shipment.ShipTo.EmailAddress));
        EmailAddress_tag.appendChild(EmailAddress_text);
        ShipTo_tag.appendChild(EmailAddress_tag);
        
        TaxIDNumber_tag = docNode.createElement('TaxIDNumber');
        TaxIDNumber_text = docNode.createTextNode(char(Shipment.ShipTo.TaxIDNumber));
        TaxIDNumber_tag.appendChild(TaxIDNumber_text);
        ShipTo_tag.appendChild(TaxIDNumber_tag);
        
        % ------------------------------------------------------------- %
        %                    ShipmentInformation                        %
        % ------------------------------------------------------------- %
        
        ShipmentInformation_tag = docNode.createElement('ShipmentInformation');
        % ShipmentInformation_text = docNode.createTextNode('ShipmentInformation_text');
        % ShipmentInformation_tag.appendChild(ShipmentInformation_text);
        OpenShipment_node.appendChild(ShipmentInformation_tag);
        
        ServiceType_tag = docNode.createElement('ServiceType');
        ServiceType_text = docNode.createTextNode(char(Shipment.ShipmentInformation.ServiceType));
        ServiceType_tag.appendChild(ServiceType_text);
        ShipmentInformation_tag.appendChild(ServiceType_tag);
        
        PackageType_tag = docNode.createElement('PackageType');
        PackageType_text = docNode.createTextNode(char(Shipment.ShipmentInformation.PackageType));
        PackageType_tag.appendChild(PackageType_text);
        ShipmentInformation_tag.appendChild(PackageType_tag);
        
        NumberOfPackages_tag = docNode.createElement('NumberOfPackages');
        NumberOfPackages_text = docNode.createTextNode(char(Shipment.ShipmentInformation.NumberOfPackages));
        NumberOfPackages_tag.appendChild(NumberOfPackages_text);
        ShipmentInformation_tag.appendChild(NumberOfPackages_tag);
        
        ShipmentActualWeight_tag = docNode.createElement('ShipmentActualWeight');
        ShipmentActualWeight_text = docNode.createTextNode(char(Shipment.ShipmentInformation.ShipmentActualWeight));
        ShipmentActualWeight_tag.appendChild(ShipmentActualWeight_text);
        ShipmentInformation_tag.appendChild(ShipmentActualWeight_tag);
        
        DescriptionOfGoods_tag = docNode.createElement('DescriptionOfGoods');
        DescriptionOfGoods_text = docNode.createTextNode(char(Shipment.ShipmentInformation.DescriptionOfGoods));
        DescriptionOfGoods_tag.appendChild(DescriptionOfGoods_text);
        ShipmentInformation_tag.appendChild(DescriptionOfGoods_tag);
        
        Reference1_tag = docNode.createElement('Reference1');
        Reference1_text = docNode.createTextNode(char(Shipment.ShipmentInformation.Reference1));
        Reference1_tag.appendChild(Reference1_text);
        ShipmentInformation_tag.appendChild(Reference1_tag);
        
        Reference2_tag = docNode.createElement('Reference2');
        Reference2_text = docNode.createTextNode(char(Shipment.ShipmentInformation.Reference2));
        Reference2_tag.appendChild(Reference2_text);
        ShipmentInformation_tag.appendChild(Reference2_tag);
        
        ShipperNumber_tag = docNode.createElement('ShipperNumber');
        ShipperNumber_text = docNode.createTextNode(char(Shipment.ShipmentInformation.ShipperNumber));
        ShipperNumber_tag.appendChild(ShipperNumber_text);
        ShipmentInformation_tag.appendChild(ShipperNumber_tag);
        
        BillingOption_tag = docNode.createElement('BillingOption');
        BillingOption_text = docNode.createTextNode(char(Shipment.ShipmentInformation.BillingOption));
        BillingOption_tag.appendChild(BillingOption_text);
        ShipmentInformation_tag.appendChild(BillingOption_tag);
        
        % ------------------------------------------------------------- %
        %                          QVNOption                            %
        % ------------------------------------------------------------- %
        
        QVNOption_tag = docNode.createElement('QVNOption');
        % ShipmentInformation_text = docNode.createTextNode('ShipmentInformation_text');
        % ShipmentInformation_tag.appendChild(ShipmentInformation_text);
        ShipmentInformation_tag.appendChild(QVNOption_tag);
        
        % ------------------------------------------------------------- %
        %               QVNRecipientAndNotificationTypes                %
        % ------------------------------------------------------------- %
        
        % Get the amount of emailaddresses out of the variable
        %emailaddresses = strrep(Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.NotificationEMailAddress,' ','');
        %rgx = '[a-z0-9_]+@[a-z0-9]+(\.[a-z0-9]+)+';
        rgx = '[a-zA-Z0-9._%''+-]+@([a-zA-Z0-9._-])+\.([a-zA-Z]{2,4})';
        emailaddresses = regexpi(Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.NotificationEMailAddress,rgx,'match');
        
        % Insert in the xml for every email address - CUSTOMERS
        for cea = 1:size(emailaddresses,2)
            cea2 = char(emailaddresses(1,cea));
            % Customer
            % --------
            QVNRecipientAndNotificationTypes_tag = docNode.createElement('QVNRecipientAndNotificationTypes');
            % ShipmentInformation_text = docNode.createTextNode('ShipmentInformation_text');
            % ShipmentInformation_tag.appendChild(ShipmentInformation_text);
            QVNOption_tag.appendChild(QVNRecipientAndNotificationTypes_tag);
            
            ContactName_tag = docNode.createElement('ContactName');
            ContactName_text = docNode.createTextNode(char(Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.ContactName));
            ContactName_tag.appendChild(ContactName_text);
            QVNRecipientAndNotificationTypes_tag.appendChild(ContactName_tag);
            
            EMailAddress2_tag = docNode.createElement('EMailAddress');
            EMailAddress2_text = docNode.createTextNode(cea2);
            EMailAddress2_tag.appendChild(EMailAddress2_text);
            QVNRecipientAndNotificationTypes_tag.appendChild(EMailAddress2_tag);
            
            Ship_tag = docNode.createElement('Ship');
            Ship_text = docNode.createTextNode(char(Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.Ship));
            Ship_tag.appendChild(Ship_text);
            QVNRecipientAndNotificationTypes_tag.appendChild(Ship_tag);
            
            Exception_tag = docNode.createElement('Exception');
            Exception_text = docNode.createTextNode(char(Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.Exception));
            Exception_tag.appendChild(Exception_text);
            QVNRecipientAndNotificationTypes_tag.appendChild(Exception_tag);
            
            Delivery_tag = docNode.createElement('Delivery');
            Delivery_text = docNode.createTextNode(char(Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.Delivery));
            Delivery_tag.appendChild(Delivery_text);
            QVNRecipientAndNotificationTypes_tag.appendChild(Delivery_tag);
        end
        
        
        % Insert in the xml for every email address - INTERNAL
        emailaddresses_internal = {'support@rsprint.be','info.motion@materialise.be'};
        contactname_internal = {'RS Print support','Info Materialise Motion'};
        for cea = 1:size(emailaddresses_internal,2)
            cea2 = char(emailaddresses_internal(1,cea));
            ccn2 = char(contactname_internal(1,cea));
            % Customer
            % --------
            QVNRecipientAndNotificationTypes_tag = docNode.createElement('QVNRecipientAndNotificationTypes');
            % ShipmentInformation_text = docNode.createTextNode('ShipmentInformation_text');
            % ShipmentInformation_tag.appendChild(ShipmentInformation_text);
            QVNOption_tag.appendChild(QVNRecipientAndNotificationTypes_tag);
            
            ContactName_tag = docNode.createElement('ContactName');
            ContactName_text = docNode.createTextNode(ccn2);
            ContactName_tag.appendChild(ContactName_text);
            QVNRecipientAndNotificationTypes_tag.appendChild(ContactName_tag);
            
            EMailAddress2_tag = docNode.createElement('EMailAddress');
            EMailAddress2_text = docNode.createTextNode(cea2);
            EMailAddress2_tag.appendChild(EMailAddress2_text);
            QVNRecipientAndNotificationTypes_tag.appendChild(EMailAddress2_tag);
            
            Ship_tag = docNode.createElement('Ship');
            Ship_text = docNode.createTextNode('1');
            Ship_tag.appendChild(Ship_text);
            QVNRecipientAndNotificationTypes_tag.appendChild(Ship_tag);
            
            Exception_tag = docNode.createElement('Exception');
            Exception_text = docNode.createTextNode('1');
            Exception_tag.appendChild(Exception_text);
            QVNRecipientAndNotificationTypes_tag.appendChild(Exception_tag);
            
            Delivery_tag = docNode.createElement('Delivery');
            Delivery_text = docNode.createTextNode('1');
            Delivery_tag.appendChild(Delivery_text);
            QVNRecipientAndNotificationTypes_tag.appendChild(Delivery_tag);
        end
        
        
        
        
%         % RS Print
%         % --------
%         QVNRecipientAndNotificationTypes_tag = docNode.createElement('QVNRecipientAndNotificationTypes');
%         % ShipmentInformation_text = docNode.createTextNode('ShipmentInformation_text');
%         % ShipmentInformation_tag.appendChild(ShipmentInformation_text);
%         QVNOption_tag.appendChild(QVNRecipientAndNotificationTypes_tag);
%         
%         ContactName_tag = docNode.createElement('ContactName');
%         ContactName_text = docNode.createTextNode('RS Print Support');
%         ContactName_tag.appendChild(ContactName_text);
%         QVNRecipientAndNotificationTypes_tag.appendChild(ContactName_tag);
%         
%         EMailAddress2_tag = docNode.createElement('EMailAddress');
%         EMailAddress2_text = docNode.createTextNode('support@rsprint.be');
%         EMailAddress2_tag.appendChild(EMailAddress2_text);
%         QVNRecipientAndNotificationTypes_tag.appendChild(EMailAddress2_tag);
%         
%         Ship_tag = docNode.createElement('Ship');
%         Ship_text = docNode.createTextNode('1');
%         Ship_tag.appendChild(Ship_text);
%         QVNRecipientAndNotificationTypes_tag.appendChild(Ship_tag);
%         
%         Exception_tag = docNode.createElement('Exception');
%         Exception_text = docNode.createTextNode('1');
%         Exception_tag.appendChild(Exception_text);
%         QVNRecipientAndNotificationTypes_tag.appendChild(Exception_tag);
%         
%         Delivery_tag = docNode.createElement('Delivery');
%         Delivery_text = docNode.createTextNode('1');
%         Delivery_tag.appendChild(Delivery_text);
%         QVNRecipientAndNotificationTypes_tag.appendChild(Delivery_tag);
        
        ShipFromCompanyOrName_tag = docNode.createElement('ShipFromCompanyOrName');
        ShipFromCompanyOrName_text = docNode.createTextNode(char(Shipment.ShipmentInformation.QVNOption.ShipFromCompanyOrName));
        ShipFromCompanyOrName_tag.appendChild(ShipFromCompanyOrName_text);
        QVNOption_tag.appendChild(ShipFromCompanyOrName_tag);
        
        FailedEMailAddress_tag = docNode.createElement('FailedEMailAddress');
        FailedEMailAddress_text = docNode.createTextNode(char(Shipment.ShipmentInformation.QVNOption.FailedEMailAddress));
        FailedEMailAddress_tag.appendChild(FailedEMailAddress_text);
        QVNOption_tag.appendChild(FailedEMailAddress_tag);
        
        SubjectLine_tag = docNode.createElement('SubjectLine');
        SubjectLine_text = docNode.createTextNode(char(Shipment.ShipmentInformation.QVNOption.SubjectLine));
        SubjectLine_tag.appendChild(SubjectLine_text);
        QVNOption_tag.appendChild(SubjectLine_tag);
        
        Memo_tag = docNode.createElement('Memo');
        Memo_text = docNode.createTextNode(char(Shipment.ShipmentInformation.QVNOption.Memo));
        Memo_tag.appendChild(Memo_text);
        QVNOption_tag.appendChild(Memo_tag);
        
        % nonEU
        if strcmp(char(UPSfile_shipment(cr,eu_col)),'Noooo') == 1 && strcmp(values.CountryOfShipment,'BE') == 1
            
            if importerpresent == 1 
                % ------------------------------------------------------------- %
                %              Importer + Third Party Receiver                  %
                % ------------------------------------------------------------- %
                
                
                % Importer

                Importer_tag = docNode.createElement('Importer');
                % ShipmentInformation_text = docNode.createTextNode('ShipmentInformation_text');
                % ShipmentInformation_tag.appendChild(ShipmentInformation_text);
                OpenShipment_node.appendChild(Importer_tag);

                CompanyOrName_tag = docNode.createElement('CompanyOrName');
                CompanyOrName_text = docNode.createTextNode(char(Shipment.Importer.CompanyOrName));
                CompanyOrName_tag.appendChild(CompanyOrName_text);
                Importer_tag.appendChild(CompanyOrName_tag);

                Attention_tag = docNode.createElement('Attention');
                Attention_text = docNode.createTextNode(char(Shipment.Importer.Attention));
                Attention_tag.appendChild(Attention_text);
                Importer_tag.appendChild(Attention_tag);

                Address1_tag = docNode.createElement('Address1');
                Address1_text = docNode.createTextNode(char(Shipment.Importer.Address1));
                Address1_tag.appendChild(Address1_text);
                Importer_tag.appendChild(Address1_tag);

                Address2_tag = docNode.createElement('Address2');
                Address2_text = docNode.createTextNode(char(Shipment.Importer.Address2));
                Address2_tag.appendChild(Address2_text);
                Importer_tag.appendChild(Address2_tag);

                Address3_tag = docNode.createElement('Address3');
                Address3_text = docNode.createTextNode(char(Shipment.Importer.Address3));
                Address3_tag.appendChild(Address3_text);
                Importer_tag.appendChild(Address3_tag);

                CityOrTown_tag = docNode.createElement('CityOrTown');
                CityOrTown_text = docNode.createTextNode(char(Shipment.Importer.CityOrTown));
                CityOrTown_tag.appendChild(CityOrTown_text);
                Importer_tag.appendChild(CityOrTown_tag);

                CountryTerritory_tag = docNode.createElement('CountryTerritory');
                CountryTerritory_text = docNode.createTextNode(char(Shipment.Importer.CountryTerritory));
                CountryTerritory_tag.appendChild(CountryTerritory_text);
                Importer_tag.appendChild(CountryTerritory_tag);

                PostalCode_tag = docNode.createElement('PostalCode');
                PostalCode_text = docNode.createTextNode(char(Shipment.Importer.PostalCode));
                PostalCode_tag.appendChild(PostalCode_text);
                Importer_tag.appendChild(PostalCode_tag);

                StateProvinceCounty_tag = docNode.createElement('StateProvinceCounty');
                StateProvinceCounty_text = docNode.createTextNode(char(Shipment.Importer.StateProvinceCounty));
                StateProvinceCounty_tag.appendChild(StateProvinceCounty_text);
                Importer_tag.appendChild(StateProvinceCounty_tag);

                Telephone_tag = docNode.createElement('Telephone');
                Telephone_text = docNode.createTextNode(char(Shipment.Importer.Telephone));
                Telephone_tag.appendChild(Telephone_text);
                Importer_tag.appendChild(Telephone_tag);

                UpsAccountNumber_tag = docNode.createElement('UpsAccountNumber');
                UpsAccountNumber_text = docNode.createTextNode(char(Shipment.Importer.UpsAccountNumber));
                UpsAccountNumber_tag.appendChild(UpsAccountNumber_text);
                Importer_tag.appendChild(UpsAccountNumber_tag);

                TaxIDNumber_tag = docNode.createElement('TaxIDNumber');
                TaxIDNumber_text = docNode.createTextNode(char(Shipment.Importer.TaxIDNumber));
                TaxIDNumber_tag.appendChild(TaxIDNumber_text);
                Importer_tag.appendChild(TaxIDNumber_tag);
                
                
                
                % Third party receiver
                % = the same information as the importer (but without the TAX ID number

                ThirdPartyReceiver_tag = docNode.createElement('ThirdPartyReceiver');
                % ShipmentInformation_text = docNode.createTextNode('ShipmentInformation_text');
                % ShipmentInformation_tag.appendChild(ShipmentInformation_text);
                OpenShipment_node.appendChild(ThirdPartyReceiver_tag);

                CompanyOrName_tag = docNode.createElement('CompanyOrName');
                CompanyOrName_text = docNode.createTextNode(char(Shipment.Importer.CompanyOrName));
                CompanyOrName_tag.appendChild(CompanyOrName_text);
                ThirdPartyReceiver_tag.appendChild(CompanyOrName_tag);

                Attention_tag = docNode.createElement('Attention');
                Attention_text = docNode.createTextNode(char(Shipment.Importer.Attention));
                Attention_tag.appendChild(Attention_text);
                ThirdPartyReceiver_tag.appendChild(Attention_tag);

                Address1_tag = docNode.createElement('Address1');
                Address1_text = docNode.createTextNode(char(Shipment.Importer.Address1));
                Address1_tag.appendChild(Address1_text);
                ThirdPartyReceiver_tag.appendChild(Address1_tag);

                Address2_tag = docNode.createElement('Address2');
                Address2_text = docNode.createTextNode(char(Shipment.Importer.Address2));
                Address2_tag.appendChild(Address2_text);
                ThirdPartyReceiver_tag.appendChild(Address2_tag);

                Address3_tag = docNode.createElement('Address3');
                Address3_text = docNode.createTextNode(char(Shipment.Importer.Address3));
                Address3_tag.appendChild(Address3_text);
                ThirdPartyReceiver_tag.appendChild(Address3_tag);

                CityOrTown_tag = docNode.createElement('CityOrTown');
                CityOrTown_text = docNode.createTextNode(char(Shipment.Importer.CityOrTown));
                CityOrTown_tag.appendChild(CityOrTown_text);
                ThirdPartyReceiver_tag.appendChild(CityOrTown_tag);

                CountryTerritory_tag = docNode.createElement('CountryTerritory');
                CountryTerritory_text = docNode.createTextNode(char(Shipment.Importer.CountryTerritory));
                CountryTerritory_tag.appendChild(CountryTerritory_text);
                ThirdPartyReceiver_tag.appendChild(CountryTerritory_tag);

                PostalCode_tag = docNode.createElement('PostalCode');
                PostalCode_text = docNode.createTextNode(char(Shipment.Importer.PostalCode));
                PostalCode_tag.appendChild(PostalCode_text);
                ThirdPartyReceiver_tag.appendChild(PostalCode_tag);

                StateProvinceCounty_tag = docNode.createElement('StateProvinceCounty');
                StateProvinceCounty_text = docNode.createTextNode(char(Shipment.Importer.StateProvinceCounty));
                StateProvinceCounty_tag.appendChild(StateProvinceCounty_text);
                ThirdPartyReceiver_tag.appendChild(StateProvinceCounty_tag);

                Telephone_tag = docNode.createElement('Telephone');
                Telephone_text = docNode.createTextNode(char(Shipment.Importer.Telephone));
                Telephone_tag.appendChild(Telephone_text);
                ThirdPartyReceiver_tag.appendChild(Telephone_tag);

                UpsAccountNumber_tag = docNode.createElement('UpsAccountNumber');
                UpsAccountNumber_text = docNode.createTextNode(char(Shipment.Importer.UpsAccountNumber));
                UpsAccountNumber_tag.appendChild(UpsAccountNumber_text);
                ThirdPartyReceiver_tag.appendChild(UpsAccountNumber_tag);

            end
            
            % ------------------------------------------------------------- %
            %                    ShipmentInformation                        %
            % ------------------------------------------------------------- %
            
            ProcessAsPaperless_tag = docNode.createElement('ProcessAsPaperless');
            ProcessAsPaperless_text = docNode.createTextNode(char(Shipment.ShipmentInformation.ProcessAsPaperless));
            ProcessAsPaperless_tag.appendChild(ProcessAsPaperless_text);
            ShipmentInformation_tag.appendChild(ProcessAsPaperless_tag);
            
            BillTransportationTo_tag = docNode.createElement('BillTransportationTo');
            BillTransportationTo_text = docNode.createTextNode(char(Shipment.ShipmentInformation.BillTransportationTo));
            BillTransportationTo_tag.appendChild(BillTransportationTo_text);
            ShipmentInformation_tag.appendChild(BillTransportationTo_tag);
            
            BillDutyTaxTo_tag = docNode.createElement('BillDutyTaxTo');
            BillDutyTaxTo_text = docNode.createTextNode(char(Shipment.ShipmentInformation.BillDutyTaxTo));
            BillDutyTaxTo_tag.appendChild(BillDutyTaxTo_text);
            ShipmentInformation_tag.appendChild(BillDutyTaxTo_tag);
            
            % ------------------------------------------------------------- %
            %                InternationalDocumentation_tag                 %
            % ------------------------------------------------------------- %
            
            InternationalDocumentation_tag = docNode.createElement('InternationalDocumentation');
            % ShipmentInformation_text = docNode.createTextNode('ShipmentInformation_text');
            % ShipmentInformation_tag.appendChild(ShipmentInformation_text);
            OpenShipment_node.appendChild(InternationalDocumentation_tag);
            
            InvoiceImporterSameAsShipTo_tag = docNode.createElement('InvoiceImporterSameAsShipTo');
            InvoiceImporterSameAsShipTo_text = docNode.createTextNode(char(Shipment.InternationalDocumentation.InvoiceImporterSameAsShipTo));
            InvoiceImporterSameAsShipTo_tag.appendChild(InvoiceImporterSameAsShipTo_text);
            InternationalDocumentation_tag.appendChild(InvoiceImporterSameAsShipTo_tag);
            
            InvoiceTermOfSale_tag = docNode.createElement('InvoiceTermOfSale');
            InvoiceTermOfSale_text = docNode.createTextNode(char(Shipment.InternationalDocumentation.InvoiceTermOfSale));
            InvoiceTermOfSale_tag.appendChild(InvoiceTermOfSale_text);
            InternationalDocumentation_tag.appendChild(InvoiceTermOfSale_tag);
            
            InvoiceReasonForExport_tag = docNode.createElement('InvoiceReasonForExport');
            InvoiceReasonForExport_text = docNode.createTextNode(char(Shipment.InternationalDocumentation.InvoiceReasonForExport));
            InvoiceReasonForExport_tag.appendChild(InvoiceReasonForExport_text);
            InternationalDocumentation_tag.appendChild(InvoiceReasonForExport_tag);
            
            InvoiceDeclarationStatement_tag = docNode.createElement('InvoiceDeclarationStatement');
            InvoiceDeclarationStatement_text = docNode.createTextNode(char(Shipment.InternationalDocumentation.InvoiceDeclarationStatement));
            InvoiceDeclarationStatement_tag.appendChild(InvoiceDeclarationStatement_text);
            InternationalDocumentation_tag.appendChild(InvoiceDeclarationStatement_tag);
            
            InvoiceAdditionalComments_tag = docNode.createElement('InvoiceAdditionalComments');
            InvoiceAdditionalComments_text = docNode.createTextNode(char(Shipment.InternationalDocumentation.InvoiceAdditionalComments));
            InvoiceAdditionalComments_tag.appendChild(InvoiceAdditionalComments_text);
            InternationalDocumentation_tag.appendChild(InvoiceAdditionalComments_tag);
            
            InvoiceCurrencyCode_tag = docNode.createElement('InvoiceCurrencyCode');
            InvoiceCurrencyCode_text = docNode.createTextNode(char(Shipment.InternationalDocumentation.InvoiceCurrencyCode));
            InvoiceCurrencyCode_tag.appendChild(InvoiceCurrencyCode_text);
            InternationalDocumentation_tag.appendChild(InvoiceCurrencyCode_tag);
            
            InvoiceLineTotal_tag = docNode.createElement('InvoiceLineTotal');
            InvoiceLineTotal_text = docNode.createTextNode(char(Shipment.InternationalDocumentation.InvoiceLineTotal));
            InvoiceLineTotal_tag.appendChild(InvoiceLineTotal_text);
            InternationalDocumentation_tag.appendChild(InvoiceLineTotal_tag);
            
            
            % ------------------------------------------------------------- %
            %                             Goods                             %
            % ------------------------------------------------------------- %
            for cg = 1:goodscounter
                
                Goods_tag = docNode.createElement('Goods');
                % ShipmentInformation_text = docNode.createTextNode('ShipmentInformation_text');
                % ShipmentInformation_tag.appendChild(ShipmentInformation_text);
                OpenShipment_node.appendChild(Goods_tag);
                
                DescriptionOfGood_tag = docNode.createElement('DescriptionOfGood');
                eval([ 'DescriptionOfGood_text = docNode.createTextNode(Shipment.Goods' num2str(cg) '.DescriptionOfGood);']);
                DescriptionOfGood_tag.appendChild(DescriptionOfGood_text);
                Goods_tag.appendChild(DescriptionOfGood_tag);
                
                Inv_NAFTA_TariffCode_tag = docNode.createElement('Inv-NAFTA-TariffCode');
                %Inv_NAFTA_TariffCode_text = docNode.createTextNode('Inv-NAFTA-TariffCode_text');
                eval([ 'Inv_NAFTA_TariffCode_text = docNode.createTextNode(Shipment.Goods' num2str(cg) '.Inv_NAFTA_TariffCode);']);
                Inv_NAFTA_TariffCode_tag.appendChild(Inv_NAFTA_TariffCode_text);
                Goods_tag.appendChild(Inv_NAFTA_TariffCode_tag);
                
                Inv_NAFTA_CO_CountryTerritoryOfOrigin_tag = docNode.createElement('Inv-NAFTA-CO-CountryTerritoryOfOrigin');
                %Inv_NAFTA_CO_CountryTerritoryOfOrigin_text = docNode.createTextNode('Inv-NAFTA-CO-CountryTerritoryOfOrigin°textt');
                eval([ 'Inv_NAFTA_CO_CountryTerritoryOfOrigin_text = docNode.createTextNode(Shipment.Goods' num2str(cg) '.Inv_NAFTA_CO_CountryTerritoryOfOrigin);']);
                Inv_NAFTA_CO_CountryTerritoryOfOrigin_tag.appendChild(Inv_NAFTA_CO_CountryTerritoryOfOrigin_text);
                Goods_tag.appendChild(Inv_NAFTA_CO_CountryTerritoryOfOrigin_tag);
                
                InvoiceUnits_tag = docNode.createElement('InvoiceUnits');
                %InvoiceUnits_text = docNode.createTextNode('InvoiceUnits_text');
                eval([ 'InvoiceUnits_text = docNode.createTextNode(Shipment.Goods' num2str(cg) '.InvoiceUnits);']);
                InvoiceUnits_tag.appendChild(InvoiceUnits_text);
                Goods_tag.appendChild(InvoiceUnits_tag);
                
                InvoiceUnitOfMeasure_tag = docNode.createElement('InvoiceUnitOfMeasure');
                %InvoiceUnitOfMeasure_text = docNode.createTextNode('InvoiceUnitOfMeasure_text');
                eval([ 'InvoiceUnitOfMeasure_text = docNode.createTextNode(Shipment.Goods' num2str(cg) '.InvoiceUnitOfMeasure);']);
                InvoiceUnitOfMeasure_tag.appendChild(InvoiceUnitOfMeasure_text);
                Goods_tag.appendChild(InvoiceUnitOfMeasure_tag);
                
                Invoice_SED_UnitPrice_tag = docNode.createElement('Invoice-SED-UnitPrice');
                %Invoice_SED_UnitPrice_text = docNode.createTextNode('Invoice-SED-UnitPrice_text');
                eval([ 'Invoice_SED_UnitPrice_text = docNode.createTextNode(Shipment.Goods' num2str(cg) '.Invoice_SED_UnitPrice);']);
                Invoice_SED_UnitPrice_tag.appendChild(Invoice_SED_UnitPrice_text);
                Goods_tag.appendChild(Invoice_SED_UnitPrice_tag);
                
                InvoiceCurrencyCode_tag = docNode.createElement('InvoiceCurrencyCode');
                %InvoiceCurrencyCode_text = docNode.createTextNode('InvoiceCurrencyCode_text');
                eval([ 'InvoiceCurrencyCode_text = docNode.createTextNode(Shipment.Goods' num2str(cg) '.InvoiceCurrencyCode);']);
                InvoiceCurrencyCode_tag.appendChild(InvoiceCurrencyCode_text);
                Goods_tag.appendChild(InvoiceCurrencyCode_tag);
                
            end
        end
        
        %docNode.appendChild(docNode.createComment('this is a comment'));
        xmlFileName = ['Output\WorldShip\' values.y values.mo values.d '_' values.h values.mi values.s  '_Shipment_' num2str(cr) '_' shipmentnumber '_' Shipment.ShipTo.CompanyOrName '.xml'];
        xmlwrite(xmlFileName,docNode);
        %type(xmlFileName);
        Shipment.ShipmentRowNumber = cr;
        content = char(UPSfile_shipment(cr,col_serialnumbers));
        Shipment.CaseIDs = readmanualcaseids({},'none',0,content,0);    
        save(['Temp\shipment_' shipmentnumber '.mat'],'Shipment');
        clear Shipment
    elseif strcmp(cell2mat(UPSfile_shipment(cr,col_shipped)),'-') == 1 && strcmp(UPSfile_shipment(cr,col_service),'UPS') == 1 ...
            && hasgotshipnr == 0 && strcmp(char(UPSfile_shipment(cr,col_ShippedFrom)),values.CountryOfShipment) == 1
        load(['Temp\shipment_' shipmentnumber '.mat'],'Shipment');
        Shipment.Subdeliveries.(cstn_c_).CompanyNumber = cstn_c_;
        Shipment.Subdeliveries.(cstn_c_).CompanyName = customers.(cstn_c_).Company;
        Shipment.Subdeliveries.(cstn_c_).ShipmentRowNumber = cr;
        content = char(UPSfile_shipment(cr,col_serialnumbers));
        Shipment.Subdeliveries.(cstn_c_).CaseIDs = readmanualcaseids({},'none',0,content,0);     
        save(['Temp\shipment_' shipmentnumber '.mat'],'Shipment');
        clear Shipment
    end
end

save Temp\list_shipments.mat list_shipments

end