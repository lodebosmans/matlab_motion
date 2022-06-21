function createnavisionpurchaseorderinput(values)

load Temp\UPSfile_shipment.mat UPSfile_shipment
load Temp\UPSfile_product.mat UPSfile_product %#ok<NASGU>
load Temp\customers customers 

disp('Creating WorldShip input - please wait')

[nrofrows,nrofcols] = size(UPSfile_shipment); %#ok<ASGLU>

col.shipmentlabel = UPSfile_shipment(1,:);

col.productlabel = {'Insoles','Demo insoles','Footscan pressure plate 0.5m','Footscan pressure plate 1.0m',...
                    'Footscan pressure plate 1.5m','Footscan pressure plate 2.0m','iQube mini','iQube','Tiger',...
                    'Marketing Materials', 'Flightcase 0.5m','Flightcase 1.0m','Flightcase 1.5m','Flightcase 2.0m', ...
                    'Runways thin','Runways thick','2D interface box','3D interface box'};
                
shipmentcounter = 0;

for cr = 2:nrofrows
    % Check is the shipment in the row is already shipped. If not, fetch all necessary data
    col_shipped = catchcolumnindex({'Shipped'},col.shipmentlabel,1);
    col_shipped = cell2mat(col_shipped(2,1));
    col_service = catchcolumnindex({'Service'},col.shipmentlabel,1);
    col_service = cell2mat(col_service(2,1));
    col_shipnr = catchcolumnindex({'ShipmentNumber'},col.shipmentlabel,1);
    col_shipnr = cell2mat(col_shipnr(2,1));
    
    % Fetch the data
    if isempty(cell2mat(UPSfile_shipment(cr,col_shipped))) == 1 && strcmp(UPSfile_shipment(cr,col_service),'UPS') == 1 && isempty(cell2mat(UPSfile_shipment(cr,col_shipnr))) == 0
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
            clear all
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
                       
            % Goods  
            % -----
            insole_col = catchcolumnindex({'Insole'},col.shipmentlabel,1);
            insole_col = cell2mat(insole_col(2,1));
            d3Box_col = catchcolumnindex({'3DBox'},col.shipmentlabel,1);
            d3Box_col = cell2mat(d3Box_col(2,1));
            difference = abs(1 - insole_col);
            fs05_col = catchcolumnindex({'FS05'},col.shipmentlabel,1);
            fs05_col = cell2mat(fs05_col(2,1)); %#ok<NASGU>
            fs10_col = catchcolumnindex({'FS10'},col.shipmentlabel,1);
            fs10_col = cell2mat(fs10_col(2,1)); %#ok<NASGU>
            fs15_col = catchcolumnindex({'FS15'},col.shipmentlabel,1);
            fs15_col = cell2mat(fs15_col(2,1)); %#ok<NASGU>
            fs20_col = catchcolumnindex({'FS20'},col.shipmentlabel,1);
            fs20_col = cell2mat(fs20_col(2,1)); %#ok<NASGU>
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
        
        % Write to file
        
        docNode = com.mathworks.xml.XMLUtils.createDocument('SalesOrder');
        docRootNode = docNode.getDocumentElement;
        
        % ------------------------------------------------------------- %
        %                       SalesHeader                             %
        % ------------------------------------------------------------- %
        
        SalesHeader_tag = docNode.createElement('SalesHeader');
        docRootNode.appendChild(SalesHeader_tag);
        
        SellToCustomerNo_tag = docNode.createElement('SellToCustomerNo');
        SellToCustomerNo_text = docNode.createTextNode('TEST');
        SellToCustomerNo_tag.appendChild(SellToCustomerNo_text);
        SalesHeader_tag.appendChild(SellToCustomerNo_tag);
        
        CustomerReference_tag = docNode.createElement('CustomerReference');
        CustomerReference_text = docNode.createTextNode('TEST');
        CustomerReference_tag.appendChild(CustomerReference_text);
        SalesHeader_tag.appendChild(CustomerReference_tag);
        
        RequestedDeliveryDate_tag = docNode.createElement('RequestedDeliveryDate');
        RequestedDeliveryDate_text = docNode.createTextNode('TEST');
        RequestedDeliveryDate_tag.appendChild(RequestedDeliveryDate_text);
        SalesHeader_tag.appendChild(RequestedDeliveryDate_tag);
        
        CaseID_tag = docNode.createElement('CaseID');
        CaseID_text = docNode.createTextNode('TEST');
        CaseID_tag.appendChild(CaseID_text);
        SalesHeader_tag.appendChild(CaseID_tag);
        
        SalesPersonCode_tag = docNode.createElement('SalesPersonCode');
        SalesPersonCode_text = docNode.createTextNode('TEST');
        SalesPersonCode_tag.appendChild(SalesPersonCode_text);
        SalesHeader_tag.appendChild(SalesPersonCode_tag);
        
        Location_tag = docNode.createElement('Location');
        Location_text = docNode.createTextNode('TEST');
        Location_tag.appendChild(Location_text);
        SalesHeader_tag.appendChild(Location_tag);
        
        ShipToCode_tag = docNode.createElement('ShipToCode');
        ShipToCode_text = docNode.createTextNode('TEST');
        ShipToCode_tag.appendChild(ShipToCode_text);
        SalesHeader_tag.appendChild(ShipToCode_tag);
        
        ShipmentMethodCode_tag = docNode.createElement('ShipmentMethodCode');
        ShipmentMethodCode_text = docNode.createTextNode('TEST');
        ShipmentMethodCode_tag.appendChild(ShipmentMethodCode_text);
        SalesHeader_tag.appendChild(ShipmentMethodCode_tag);
        
        ShippingAgentCode_tag = docNode.createElement('ShippingAgentCode');
        ShippingAgentCode_text = docNode.createTextNode('TEST');
        ShippingAgentCode_tag.appendChild(ShippingAgentCode_text);
        SalesHeader_tag.appendChild(ShippingAgentCode_tag);
        
        ShippingAgentServiceCode_tag = docNode.createElement('ShippingAgentServiceCode');
        ShippingAgentServiceCode_text = docNode.createTextNode('TEST');
        ShippingAgentServiceCode_tag.appendChild(ShippingAgentServiceCode_text);
        SalesHeader_tag.appendChild(ShippingAgentServiceCode_tag);
        
            % ------------------------------------------------------------- %
            %                        SalesLine                              %
            % ------------------------------------------------------------- %

            SalesLine_tag = docNode.createElement('SalesLine');
            SalesHeader_tag.appendChild(SalesLine_tag);

            Type_tag = docNode.createElement('Type');
            Type_text = docNode.createTextNode('TEST');
            Type_tag.appendChild(Type_text);
            SalesLine_tag.appendChild(Type_tag);

            ItemNo_tag = docNode.createElement('ItemNo');
            ItemNo_text = docNode.createTextNode('TEST');
            ItemNo_tag.appendChild(ItemNo_text);
            SalesLine_tag.appendChild(ItemNo_tag);

            PurchasingCode_tag = docNode.createElement('PurchasingCode');
            PurchasingCode_text = docNode.createTextNode('TEST');
            PurchasingCode_tag.appendChild(PurchasingCode_text);
            SalesLine_tag.appendChild(PurchasingCode_tag);

            Description2_tag = docNode.createElement('Description2');
            Description2_text = docNode.createTextNode('TEST');
            Description2_tag.appendChild(Description2_text);
            SalesLine_tag.appendChild(Description2_tag);

            LocationCode_tag = docNode.createElement('LocationCode');
            LocationCode_text = docNode.createTextNode('TEST');
            LocationCode_tag.appendChild(LocationCode_text);
            SalesLine_tag.appendChild(LocationCode_tag);

            Quantity_tag = docNode.createElement('Quantity');
            Quantity_text = docNode.createTextNode('TEST');
            Quantity_tag.appendChild(Quantity_text);
            SalesLine_tag.appendChild(Quantity_tag);

            UnitOfMeasureCode_tag = docNode.createElement('UnitOfMeasureCode');
            UnitOfMeasureCode_text = docNode.createTextNode('TEST');
            UnitOfMeasureCode_tag.appendChild(UnitOfMeasureCode_text);
            SalesLine_tag.appendChild(UnitOfMeasureCode_tag);

            UnitPriceExclVAT_tag = docNode.createElement('UnitPriceExclVAT');
            UnitPriceExclVAT_text = docNode.createTextNode('TEST');
            UnitPriceExclVAT_tag.appendChild(UnitPriceExclVAT_text);
            SalesLine_tag.appendChild(UnitPriceExclVAT_tag);

            LineDiscount_tag = docNode.createElement('LineDiscount');
            LineDiscount_text = docNode.createTextNode('TEST');
            LineDiscount_tag.appendChild(LineDiscount_text);
            SalesLine_tag.appendChild(LineDiscount_tag);

            LineDiscountAmount_tag = docNode.createElement('LineDiscountAmount');
            LineDiscountAmount_text = docNode.createTextNode('TEST');
            LineDiscountAmount_tag.appendChild(LineDiscountAmount_text);
            SalesLine_tag.appendChild(LineDiscountAmount_tag);
        
        
                % ------------------------------------------------------------- %
                %                    ItemTrackingLine                           %
                % ------------------------------------------------------------- %

                ItemTrackingLine_tag = docNode.createElement('ItemTrackingLine');
                SalesLine_tag.appendChild(ItemTrackingLine_tag);

                SerialNo_tag = docNode.createElement('SerialNo');
                SerialNo_text = docNode.createTextNode('TEST');
                SerialNo_tag.appendChild(SerialNo_text);
                ItemTrackingLine_tag.appendChild(SerialNo_tag);
        
        

        
        docNode.appendChild(docNode.createComment('this is a comment'));
        xmlFileName = ['Output\Navision\PurchaseOrder' num2str(shipmentcounter) '.xml'];
        xmlwrite(xmlFileName,docNode);
        type(xmlFileName);       
        
        %if values.CreateWorldShipDeliveryNote == 1
            
        %end
        
    end
end







end