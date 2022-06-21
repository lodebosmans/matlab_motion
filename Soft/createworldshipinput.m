function createworldshipinput(values)

load Temp\UPSfile_shipment.mat UPSfile_shipment
load Temp\UPSfile_product.mat UPSfile_product %#ok<NASGU>

disp('Creating WorldShip input - please wait')

[nrofrows,nrofcols] = size(UPSfile_shipment); %#ok<ASGLU>

col.shipmentlabel = UPSfile_shipment(1,:);

% col.shipmentlabel = {'Year','Month','Day','ShipmentNumber','Service',... %1-5
%                      'CustomerNumber','Trackingnumber','Weight','Dimensions','NrOfPairs',... %6-10
%                       'DemoPromo','Ready','Shipped','SerialNumbers','ContentDescription',... %11-15
%                       'Customer','Reference','ServiceCode','PackageType','BillingOption',... %16-20
%                       'NrPackages','EU','Insole','Demo','FS05',... %21-25
%                       'FS10','FS15','FS20','iQubeMini','iQube',... %26-30
%                       'Tiger','Marketing','Flightcase05','Flightcase10','Flightcase15',... %31-35
%                       'Flightcase20','RunwaysThin','RunwaysThick','2DBox','3DBox',... %36-40
%                       'InvoiceTotal','CountryCode','StateCode','PositionCustomer','InvoiceCompany',... %41-45
%                       'InvoiceContactPerson','InvoiceAddress1','InvoiceAddress2','InvoiceAddress3','InvoicePostalCode',... %46-50
%                       'InvoiceCity','InvoiceState','InvoiceCountry','InvoicePhone','DeliveryFacilityName',... %51-55
%                       'DeliveryContactPerson','DeliveryAddress1','DeliveryAddress2','DeliveryAddress3','DeliveryPostalCode',... %56-60
%                       'DeliveryCity','DeliveryState','DeliveryCountry','DeliveryPhone','NotificationEmail',... %61-65
%                       'QuantumView','OurEmail','BillTransportationTo','DutyTax','Extra1',... %66-70
%                       'Extra2','Extra3','Extra4','Extra5'}; %71-74
col.productlabel = {'Insoles','Demo insoles','Footscan pressure plate 0.5m','Footscan pressure plate 1.0m',...
                    'Footscan pressure plate 1.5m','Footscan pressure plate 2.0m','iQube mini','iQube','Tiger',...
                    'Marketing Materials', 'Flightcase 0.5m','Flightcase 1.0m','Flightcase 1.5m','Flightcase 2.0m', ...
                    'Runways thin','Runways thick','2D interface box','3D interface box'};
                
shipmentcounter = 0;

col_shipped = catchcolumnindex({'Shipped'},col.shipmentlabel,1);
col_shipped = cell2mat(col_shipped(2,1));
col_service = catchcolumnindex({'Service'},col.shipmentlabel,1);
col_service = cell2mat(col_service(2,1));
col_shipnr = catchcolumnindex({'ShipmentNumber'},col.shipmentlabel,1);
col_shipnr = cell2mat(col_shipnr(2,1));
insole_col = catchcolumnindex({'Insole'},col.shipmentlabel,1);
insole_col = cell2mat(insole_col(2,1));
d3Box_col = catchcolumnindex({'3DBox'},col.shipmentlabel,1);
d3Box_col = cell2mat(d3Box_col(2,1));
difference = abs(1 - insole_col);
fs05_col = catchcolumnindex({'FS05'},col.shipmentlabel,1);
fs05_col = cell2mat(fs05_col(2,1));
fs10_col = catchcolumnindex({'FS10'},col.shipmentlabel,1);
fs10_col = cell2mat(fs10_col(2,1));
fs15_col = catchcolumnindex({'FS15'},col.shipmentlabel,1);
fs15_col = cell2mat(fs15_col(2,1));
fs20_col = catchcolumnindex({'FS20'},col.shipmentlabel,1);
fs20_col = cell2mat(fs20_col(2,1));


for cr = 2:nrofrows
    % Check is the shipment in the row is already shipped. If not, fetch all necessary data
    %if isempty(cell2mat(UPSfile_shipment(cr,col_shipped))) == 1 && strcmp(UPSfile_shipment(cr,col_service),'UPS') == 1 && isempty(cell2mat(UPSfile_shipment(cr,col_shipnr))) == 0
    if strcmp(cell2mat(UPSfile_shipment(cr,col_shipped)),'-') == 1 && strcmp(UPSfile_shipment(cr,col_service),'UPS') == 1 && isempty(cell2mat(UPSfile_shipment(cr,col_shipnr))) == 0
    
        % Fetch the data
        shipmentcounter = shipmentcounter + 1;
        
        % Prepare all the data
        
        % ShipTo
        % ------
        temp = catchcolumnindex({'Customer'},col.shipmentlabel,1);
        temp = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = strrep(temp,'&','and');
        Shipment.ShipTo.CompanyOrName = temp;
        disp(['Currently processing row ' num2str(cr) ' for shipment ' num2str(cell2mat(UPSfile_shipment(cr,4))) ' to ' char(Shipment.ShipTo.CompanyOrName) ' (' char(UPSfile_shipment(cr,6)) ')' ]);
        temp = catchcolumnindex({'DeliveryContactPerson'},col.shipmentlabel,1);
        Shipment.ShipTo.Attention = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = catchcolumnindex({'DeliveryAddress1'},col.shipmentlabel,1);
        Shipment.ShipTo.Address1 = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = catchcolumnindex({'DeliveryAddress2'},col.shipmentlabel,1);
        Shipment.ShipTo.Address2 = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = catchcolumnindex({'DeliveryAddress3'},col.shipmentlabel,1);
        Shipment.ShipTo.Address3 = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = catchcolumnindex({'CountryCode'},col.shipmentlabel,1);
        Shipment.ShipTo.CountryTerritory = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = catchcolumnindex({'DeliveryPostalCode'},col.shipmentlabel,1);
        if isstring(UPSfile_shipment(cr,cell2mat(temp(2,1)))) == 1
            Shipment.ShipTo.PostalCode = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        else
            Shipment.ShipTo.PostalCode = num2str(cell2mat(UPSfile_shipment(cr,cell2mat(temp(2,1)))));
        end
        temp = catchcolumnindex({'DeliveryCity'},col.shipmentlabel,1);
        Shipment.ShipTo.CityOrTown = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = catchcolumnindex({'DeliveryState'},col.shipmentlabel,1);
        Shipment.ShipTo.StateProvinceCounty = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = catchcolumnindex({'DeliveryPhone'},col.shipmentlabel,1);
        if isnumeric(UPSfile_shipment(cr,cell2mat(temp(2,1))))
            Shipment.ShipTo.Telephone = num2str(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        elseif iscell(UPSfile_shipment(cr,cell2mat(temp(2,1))))
            Shipment.ShipTo.Telephone = num2str(cell2mat(UPSfile_shipment(cr,cell2mat(temp(2,1)))));
        elseif isstring(UPSfile_shipment(cr,cell2mat(temp(2,1))))
            Shipment.ShipTo.Telephone = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        else
            clear all %#ok<CLALL>
        end
        %Shipment.ShipTo.Telephone = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        %temp = catchcolumnindex({'FaxNumber'},col.shipmentlabel,1);
        Shipment.ShipTo.FaxNumber = '';
        temp = catchcolumnindex({'NotificationEmail'},col.shipmentlabel,1);
        Shipment.ShipTo.EmailAddress = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = catchcolumnindex({'TaxIDNumber'},col.shipmentlabel,1);        
        if isstring(UPSfile_shipment(cr,cell2mat(temp(2,1)))) == 1
            Shipment.ShipTo.TaxIDNumber = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        else
            Shipment.ShipTo.TaxIDNumber = num2str(cell2mat(UPSfile_shipment(cr,cell2mat(temp(2,1)))));
        end        
        %Shipment.ShipTo.TaxIDNumber = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        
        % Importer
        % --------
        % See nonEU section for the rest
        
        % ShipmentInformation
        % -------------------
        temp = catchcolumnindex({'ServiceCode'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.ServiceType = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = catchcolumnindex({'PackageType'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.PackageType = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = catchcolumnindex({'NrPackages'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.NumberOfPackages = num2str(cell2mat(UPSfile_shipment(cr,cell2mat(temp(2,1)))));
        temp = catchcolumnindex({'Weight'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.ShipmentActualWeight = num2str(cell2mat(UPSfile_shipment(cr,cell2mat(temp(2,1)))));
        temp = catchcolumnindex({'ContentDescription'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.DescriptionOfGoods = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = catchcolumnindex({'Reference'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.Reference1 = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        %temp = catchcolumnindex({'Reference2'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.Reference2 = '';
        %temp = catchcolumnindex({'ShipperNumber'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.ShipperNumber = '';
        temp = catchcolumnindex({'BillingOption'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.BillingOption = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        %temp = catchcolumnindex({'ProcessAsPaperless'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.ProcessAsPaperless = 'Y';   
        temp = catchcolumnindex({'BillTransportationTo'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.BillTransportationTo = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = catchcolumnindex({'DutyTax'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.BillDutyTaxTo = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
                
        % QVNOption
        % ---------
        
        % QVNRecipientAndNotificationTypes
        % --------------------------------
        temp = catchcolumnindex({'DeliveryContactPerson'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.ContactName = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = catchcolumnindex({'NotificationEmail'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.EMailAddress = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        %temp = catchcolumnindex({'Ship'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.Ship = '1';
        %temp = catchcolumnindex({'Exception'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.Exception = '0';
        %temp = catchcolumnindex({'Delivery'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.Delivery = '0';
        
        % QVNOption (part 2)
        % ------------------
        
        %temp = catchcolumnindex({'RS Print'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.QVNOption.ShipFromCompanyOrName = 'RS Print';
        %temp = catchcolumnindex({'support@rsprint.be'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.QVNOption.FailedEMailAddress = 'support@rsprint.be';
        temp = catchcolumnindex({'Reference'},col.shipmentlabel,1);
        Shipment.ShipmentInformation.QVNOption.SubjectLine = ['RS Print shipment: ' char(UPSfile_shipment(cr,cell2mat(temp(2,1))));];
        temp = catchcolumnindex({'Memo'},col.shipmentlabel,1);
        temp = char(UPSfile_shipment(cr,cell2mat(temp(2,1))));
        temp = strrep(temp,'   ',''); % Reduce the length of the string
        temp = strrep(temp,'-',''); % Idem ditto
        Shipment.ShipmentInformation.QVNOption.Memo = temp;
        
        eu_col = catchcolumnindex({'EU'},col.shipmentlabel,1);
        eu_col = cell2mat(eu_col(2,1));
        if strcmp(char(UPSfile_shipment(cr,eu_col)),'Noooo') == 1 || strcmp(values.CountryOfShipment,'BE') == 1
            
            % Importer
            % --------
            Shipment.Importer.CompanyOrName = 'Materialise USA';
            Shipment.Importer.Attention = 'Juston Boland';
            Shipment.Importer.Address1 = '44650 Helm Court';
            Shipment.Importer.CityOrTown = 'Plymouth';
            Shipment.Importer.CountryTerritory = 'US';
            Shipment.Importer.PostalCode = '48170';
            Shipment.Importer.StateProvinceCounty = 'MI';
            Shipment.Importer.Telephone = '0017342596445';
            Shipment.Importer.UpsAccountNumber = 'A66F38';
            Shipment.Importer.TaxIDNumber = '38-3275861';
            
            % International documentation
            % ---------------------------
            %temp = catchcolumnindex({'InvoiceImporterSameAsShipTo'},col.shipmentlabel,1);
            Shipment.InternationalDocumentation.InvoiceImporterSameAsShipTo = 'N';
            %temp = catchcolumnindex({'InvoiceTermOfSale'},col.shipmentlabel,1);
            Shipment.InternationalDocumentation.InvoiceTermOfSale = '';
            %temp = catchcolumnindex({'InvoiceReasonForExport'},col.shipmentlabel,1);
            Shipment.InternationalDocumentation.InvoiceReasonForExport = '02';
            %temp = catchcolumnindex({'InvoiceDeclarationStatement'},col.shipmentlabel,1);
            Shipment.InternationalDocumentation.InvoiceDeclarationStatement = 'I hereby certify that the information on this invoice is true and correct and the contents and value of this shipment is as stated above.';
            %temp = catchcolumnindex({'InvoiceAdditionalComments'},col.shipmentlabel,1);
            Shipment.InternationalDocumentation.InvoiceAdditionalComments = '';
            %temp = catchcolumnindex({'InvoiceCurrencyCode'},col.shipmentlabel,1);
            Shipment.InternationalDocumentation.InvoiceCurrencyCode = 'EUR';
            temp = catchcolumnindex({'InvoiceTotal'},col.shipmentlabel,1);
            Shipment.InternationalDocumentation.InvoiceLineTotal = num2str(cell2mat(UPSfile_shipment(cr,cell2mat(temp(2,1)))));
                   
            
% ----------------------------------------------------------------------------------------------------
% ----------------------------------------------------------------------------------------------------
% ----------------------------------------------------------------------------------------------------
% ----------------------------------------------------------------------------------------------------
            
            
            
            % Goods  
            % -----

            goodscounter = 0;
            for cg = insole_col:2:d3Box_col
                amount = cell2mat(UPSfile_shipment(cr,cg));
                if isnumeric(amount) == 1 && amount > 0
                    goodscounter = goodscounter + 1;
                    cp = ((cg - difference)+1)/2; %#ok<NASGU>
                    paperless = 'Y'; %#ok<NASGU>
                    createinvoice = 'Y'; %#ok<NASGU>
                    unittype = 'PC'; %#ok<NASGU>
                    currency = 'EUR'; %#ok<NASGU>
                    countryoforigin = 'BE'; %#ok<NASGU>
                    eval(['Shipment.Goods' num2str(goodscounter) '.DescriptionOfGood = char(col.productlabel(1,cp));']);
                    eval(['Shipment.Goods' num2str(goodscounter) '.Inv_NAFTA_TariffCode = num2str(cell2mat((UPSfile_product(3,cp))));']);
                    eval(['Shipment.Goods' num2str(goodscounter) '.Inv_NAFTA_CO_CountryTerritoryOfOrigin = countryoforigin;']);
                    eval(['Shipment.Goods' num2str(goodscounter) '.InvoiceUnits = num2str(amount);']);
                    eval(['Shipment.Goods' num2str(goodscounter) '.InvoiceUnitOfMeasure = unittype;']);
                    eval(['Shipment.Goods' num2str(goodscounter) '.Invoice_SED_UnitPrice = num2str(cell2mat((UPSfile_product(2,cp))));']);
                    if strcmp(Shipment.ShipmentInformation.DescriptionOfGoods,'Insoles') && strcmp(Shipment.ShipTo.CompanyOrName,'Alchemaker')
                        eval(['Shipment.Goods' num2str(goodscounter) '.Invoice_SED_UnitPrice = ''90'';']);
                        Shipment.InternationalDocumentation.InvoiceLineTotal = num2str(amount * 90);
                    end
                    eval(['Shipment.Goods' num2str(goodscounter) '.InvoiceCurrencyCode = currency;']);
                    % Print declaration of conformity for footscan plate
                    if  (cg == fs05_col || cg == fs10_col || cg == fs15_col || cg == fs20_col ) && amount > 0
                        % Specify file name
                        file = 'C:\Users\mathlab\Documents\MATLAB\Interface\Input\Templates\DeclarationOfConformityFootscan_v2.xlsx'; % This must be full path name
                        % Open Excel Automation server
                        Excel = actxserver('Excel.Application');
                        Workbooks = Excel.Workbooks;
                        % Make Excel visible or not
                        Excel.Visible=0;
                        % Open Excel file
                        Workbook = Workbooks.Open(file); %#ok<NASGU>
                        Excel.ActiveWorkbook.PrintOut(1,1,1,'False','NPI50FDAA (HP LaserJet M506)');
                        % Close Excel and clean up
                        invoke(Excel,'Quit');
                        delete(Excel);
                        clear Excel;
                        %declarationalreadyprinted = 1;
                    end
                end
            end
            clear cg
        end
        
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
        StateProvinceCounty_text = docNode.createTextNode(char(Shipment.ShipTo.StateProvinceCounty));
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
            EMailAddress2_text = docNode.createTextNode(char(Shipment.ShipmentInformation.QVNOption.QVNRecipientAndNotificationTypes.EMailAddress));
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
            
            % RS Print     
            % --------
            QVNRecipientAndNotificationTypes_tag = docNode.createElement('QVNRecipientAndNotificationTypes');
            % ShipmentInformation_text = docNode.createTextNode('ShipmentInformation_text');
            % ShipmentInformation_tag.appendChild(ShipmentInformation_text);
            QVNOption_tag.appendChild(QVNRecipientAndNotificationTypes_tag);

            ContactName_tag = docNode.createElement('ContactName');
            ContactName_text = docNode.createTextNode('RS Print Support');
            ContactName_tag.appendChild(ContactName_text);
            QVNRecipientAndNotificationTypes_tag.appendChild(ContactName_tag);

            EMailAddress2_tag = docNode.createElement('EMailAddress');
            EMailAddress2_text = docNode.createTextNode('support@rsprint.be');
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
        if strcmp(char(UPSfile_shipment(cr,eu_col)),'Noooo') == 1
            % ------------------------------------------------------------- %
            %                          Importer                             %
            % ------------------------------------------------------------- %
            
%             Importer_tag = docNode.createElement('Importer');
%             % ShipmentInformation_text = docNode.createTextNode('ShipmentInformation_text');
%             % ShipmentInformation_tag.appendChild(ShipmentInformation_text);
%             OpenShipment_node.appendChild(Importer_tag);
%             
%             CompanyOrName_tag = docNode.createElement('CompanyOrName');
%             CompanyOrName_text = docNode.createTextNode(char(Shipment.Importer.CompanyOrName));
%             CompanyOrName_tag.appendChild(CompanyOrName_text);
%             Importer_tag.appendChild(CompanyOrName_tag);
%             
%             Attention_tag = docNode.createElement('Attention');
%             Attention_text = docNode.createTextNode(char(Shipment.Importer.Attention));
%             Attention_tag.appendChild(Attention_text);
%             Importer_tag.appendChild(Attention_tag);
%             
%             Address1_tag = docNode.createElement('Address1');
%             Address1_text = docNode.createTextNode(char(Shipment.Importer.Address1));
%             Address1_tag.appendChild(Address1_text);
%             Importer_tag.appendChild(Address1_tag);
%             
%             CityOrTown_tag = docNode.createElement('CityOrTown');
%             CityOrTown_text = docNode.createTextNode(char(Shipment.Importer.CityOrTown));
%             CityOrTown_tag.appendChild(CityOrTown_text);
%             Importer_tag.appendChild(CityOrTown_tag);
%             
%             CountryTerritory_tag = docNode.createElement('CountryTerritory');
%             CountryTerritory_text = docNode.createTextNode(char(Shipment.Importer.CountryTerritory));
%             CountryTerritory_tag.appendChild(CountryTerritory_text);
%             Importer_tag.appendChild(CountryTerritory_tag);
%             
%             PostalCode_tag = docNode.createElement('PostalCode');
%             PostalCode_text = docNode.createTextNode(char(Shipment.Importer.PostalCode));
%             PostalCode_tag.appendChild(PostalCode_text);
%             Importer_tag.appendChild(PostalCode_tag);
%             
%             StateProvinceCounty_tag = docNode.createElement('StateProvinceCounty');
%             StateProvinceCounty_text = docNode.createTextNode(char(Shipment.Importer.StateProvinceCounty));
%             StateProvinceCounty_tag.appendChild(StateProvinceCounty_text);
%             Importer_tag.appendChild(StateProvinceCounty_tag);
%             
%             Telephone_tag = docNode.createElement('Telephone');
%             Telephone_text = docNode.createTextNode(char(Shipment.Importer.Telephone));
%             Telephone_tag.appendChild(Telephone_text);
%             Importer_tag.appendChild(Telephone_tag);
%             
%             UpsAccountNumber_tag = docNode.createElement('UpsAccountNumber');
%             UpsAccountNumber_text = docNode.createTextNode(char(Shipment.Importer.UpsAccountNumber));
%             UpsAccountNumber_tag.appendChild(UpsAccountNumber_text);
%             Importer_tag.appendChild(UpsAccountNumber_tag);
%                          
%             TaxIDNumber_tag = docNode.createElement('TaxIDNumber');
%             TaxIDNumber_text = docNode.createTextNode(char(Shipment.Importer.TaxIDNumber));
%             TaxIDNumber_tag.appendChild(TaxIDNumber_text);
%             Importer_tag.appendChild(TaxIDNumber_tag);            
            
            
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
        
        docNode.appendChild(docNode.createComment('this is a comment'));
        xmlFileName = ['Output\WorldShip\Shipment' num2str(shipmentcounter) '.xml'];
        xmlwrite(xmlFileName,docNode);
        %type(xmlFileName);       
        
        %if values.CreateWorldShipDeliveryNote == 1
            
        %end
        
    end
end


end