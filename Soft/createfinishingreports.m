function createfinishingreports()

load Temp\caseids_production_tocreate.mat caseids_production_tocreate

if size(caseids_production_tocreate,1) > 0
    
    load Temp\values.mat values
    load Temp\db_production.mat db_production
    %load Temp\caseids_production.mat caseids_production
    
    % Check is some of these finishing reports are already printed.
    load([values.baseversion 'Input\Variables\fr.mat'],'fr') % Always store in the interface folder.
    alert = 4;
    for cfr = 1:size(db_production.overview,1)
        string = char(db_production.overview(cfr,1));
        idx = strcmp(string,fr(:,1));
        temp = num2cell(find(idx==1));
        newinput = [values.d '/' values.mo '/20' values.y ' - ' values.Username_str];
        % if not empty, add to list of alerts
        if isempty(temp) == 0
            if alert == 4
                Excel = actxserver('Excel.Application');
                % Set it to visible
                set(Excel,'Visible',0);
                % Add a Workbook
                Workbooks = Excel.Workbooks;
                Workbook = invoke(Workbooks, 'Add');
                % Get a handle to Sheets and select Sheet 1
                Sheets = Excel.ActiveWorkBook.Sheets;
                Sheet1 = get(Sheets, 'Item', 1);
                Sheet1.Activate;
                % Add the case to the first column
                eActivesheetRange = get(Excel.Activesheet,'Range','A1:A1');
                eActivesheetRange.Value = 'Finishing reports that have been printed before.';
                eActivesheetRange = get(Excel.Activesheet,'Range','A2:A2');
                eActivesheetRange.Value = 'Please investigate if this is a reprint or rebuild to avoid mistakes.';
                Sheet1.Columns.Item(1).columnWidth = 15;
            end
            newinput2 = [char(fr(cell2mat(temp(1,1)),2)) ' /// ' newinput];
            fr(cell2mat(temp(1,1)),2) = cellstr(newinput2);
            % Add the case to the first column
            eActivesheetRange = get(Excel.Activesheet,'Range',['A' num2str(alert) ':A' num2str(alert)]);
            eActivesheetRange.Value = string;
            % Add the date and actor to the second column
            eActivesheetRange = get(Excel.Activesheet,'Range',['B' num2str(alert) ':B' num2str(alert)]);
            eActivesheetRange.Value = newinput2;
            
            alert = alert + 1;
            % If empty, add to list
        else
            sizefr = size(fr,1);
            fr(sizefr+1,1) = cellstr(string);
            fr(sizefr+1,2) = cellstr(newinput);
        end
    end
    save([values.baseversion 'Input\Variables\fr.mat'],'fr');
    if alert > 4
        % Print and close Excel
        invoke(Workbook, 'SaveAs', [pwd '\Temp\AlertFinishingReports.xlsx']);
        if values.MenuSelectPrinter > 1
            Excel.ActiveWorkbook.PrintOut(1,1,1,'False',values.MenuSelectPrinter_str);
        end
        invoke(Excel, 'Quit');
    end
    
    nrofcases = size(db_production.overview,1);
    heelpads = {'Plantar Fasciitis','Heel Spur','Fat Pad'};
    
    for caseindex = 1:nrofcases
        displaytext = ['Creating finishing reports - please wait - processing ' num2str(caseindex) ' of ' num2str(nrofcases)];
        disp(displaytext);
        % Get the caseID from the list of caseIDs
        clear currentcaseid
        currentcaseid_long = char(db_production.overview(caseindex,1));
        currentcaseid = [currentcaseid_long(1:4) currentcaseid_long(6:8) currentcaseid_long(10:12)];
        
        % Create Excel finishing report
        ExcelOutput = cell(41,3);
        ExcelOutput(1,1) = cellstr('Order ID:');
        ExcelOutput(1,2) = cellstr(db_production.raw.(currentcaseid).caseid.full);
        
        %ExcelOutput(4,1) = cellstr('Name subject:');
        %ExcelOutput(4,2) = cellstr([db_production.raw.(currentcaseid).customerfirstname ' ' db_production.raw.(currentcaseid).customersurname]);
        
        ExcelOutput(5,1) = cellstr('Practitioner:');
        ExcelOutput(5,2) = cellstr(db_production.raw.(currentcaseid).surgeon);
        
        ExcelOutput(6,1) = cellstr('Delivery address:');
        ExcelOutput(6,2) = cellstr(db_production.raw.(currentcaseid).delivery_company);
        ExcelOutput(7,2) = cellstr(db_production.raw.(currentcaseid).delivery_street);
        if strcmp(db_production.raw.(currentcaseid).delivery_state,'<None>')
            ExcelOutput(8,2) = cellstr(db_production.raw.(currentcaseid).delivery_zip_city);
        else
            ExcelOutput(8,2) = cellstr([db_production.raw.(currentcaseid).delivery_zip_city ', ' db_production.raw.(currentcaseid).delivery_state]);
        end
        ExcelOutput(9,2) = cellstr(db_production.raw.(currentcaseid).delivery_country);
        
        ExcelOutput(10,1) = cellstr('Shipping date:');
        ExcelOutput(10,2) = cellstr(db_production.raw.(currentcaseid).estimatedshippingdate);
        
        ExcelOutput(11,1) = cellstr('Reference ID:');
        ExcelOutput(11,2) = cellstr(db_production.raw.(currentcaseid).referenceid);
        
        ExcelOutput(14,1) = cellstr('Insole type:');
        ExcelOutput(14,2) = cellstr(db_production.raw.(currentcaseid).insoletype);
        
        ExcelOutput(15,1) = cellstr('Heel cup:');
        ExcelOutput(15,2) = cellstr([db_production.raw.(currentcaseid).heelcupleft ' - ' db_production.raw.(currentcaseid).heelcupright]);
        
        ExcelOutput(16,1) = cellstr('Top thickness:');
        ExcelOutput(16,2) = cellstr(db_production.raw.(currentcaseid).topthickness);
        
        ExcelOutput(17,1) = cellstr('Material:');
        ExcelOutput(17,2) = cellstr(db_production.raw.(currentcaseid).topmaterial);
        
        ExcelOutput(18,1) = cellstr('Top size:');
        wideinsole = 0;
        if strcmp(char(db_production.raw.(currentcaseid).insoletype),'Wide') == 1
            tmp1 = length(char(db_production.raw.(currentcaseid).topsize));
            tmp2 = str2double(db_production.raw.(currentcaseid).topsize(1:tmp1-5));
            ExcelOutput(18,2) = cellstr([db_production.raw.(currentcaseid).topsize(1:tmp1-3) ' + 1.0 UK = ' num2str(tmp2+1) '.0 UK*']);
            wideinsole = 1;
        else
            ExcelOutput(18,2) = cellstr(db_production.raw.(currentcaseid).topsize);
        end
        
        ExcelOutput(19,1) = cellstr('Top hardness:');
        ExcelOutput(19,2) = cellstr(db_production.raw.(currentcaseid).tophardness);
        
        ExcelOutput(20,1) = cellstr('Top layer service:');
        ExcelOutput(20,2) = cellstr(db_production.raw.(currentcaseid).servicetype);
        
        ExcelOutput(21,1) = cellstr('Heel pad - left:');
        ExcelOutput(21,2) = cellstr(db_production.raw.(currentcaseid).heelpadleft);
        
        ExcelOutput(22,1) = cellstr('Heel pad - right:');
        ExcelOutput(22,2) = cellstr(db_production.raw.(currentcaseid).heelpadright);
        
        ExcelOutput(25,1) = cellstr('Remarks:');
        ExcelOutput(25,2) = cellstr(db_production.raw.(currentcaseid).remarks);
        
        ExcelOutput(35,3) = cellstr('Materialise Motion');
        ExcelOutput(36,3) = cellstr('De Weven 7, 3583 Paal-Beringen, Belgium');
        ExcelOutput(37,3) = cellstr('+32(0) 11 36 01 79 - VAT 0551.855.071');
        ExcelOutput(38,3) = cellstr('motion@materialise.be');
        ExcelOutput(39,3) = cellstr('www.materialisemotion.com');
        
        if wideinsole == 1
            ExcelOutput(41,1) = cellstr('*Size + 1 UK in function of wide insole type');
        end
        
        %%%%%%%%%%%%%%%%%%%%
        
        % Create sample image from figure
        img = 'Input\Templates\Materialise_BL_sRGB.png';
        % Get handle to Excel COM Server
        Excel = actxserver('Excel.Application');
        % Set it to visible
        set(Excel,'Visible',0);
        % Add a Workbook
        Workbooks = Excel.Workbooks;
        Workbook = invoke(Workbooks, 'Add');
        % Get a handle to Sheets and select Sheet 1
        Sheets = Excel.ActiveWorkBook.Sheets;
        Sheet1 = get(Sheets, 'Item', 1);
        Sheet1.Activate;
        % Get a handle to Shapes for Sheet 1
        Shapes = Sheet1.Shapes;
        % Add image
        if strcmp(db_production.raw.(currentcaseid).casecancelled,'-') == 1
            Shapes.AddPicture([values.baseversion '\' img],0,1,47,1,76*1.655,76); % link to file on negative to ensure image is printable
            % dimensions =  w:1805 * h:1090
        end
        
        % Inject the data
        eActivesheetRange = get(Excel.Activesheet,'Range','A8:C48');
        eActivesheetRange.Value = ExcelOutput;
        
        % Add some exceptions
        %     if strcmp(db_production.raw.(currentcaseid).tophardness,'20 Shore')
        %         cells = Excel.ActiveSheet.Range('B24:D24');
        %         set(cells.Font, 'Bold', true);
        %         set(cells.Interior,'ColorIndex',15);
        %         if strcmp(db_production.raw.(currentcaseid).topthickness,'6 mm')
        %             cells = Excel.ActiveSheet.Range('B22:D22');
        %             set(cells.Font, 'Bold', true);
        %             set(cells.Interior,'ColorIndex',15);
        %         end
        %     end
        if strcmp(db_production.raw.(currentcaseid).heelcupleft,'Low') || strcmp(db_production.raw.(currentcaseid).heelcupright,'Low')
            cells = Excel.ActiveSheet.Range('B22:D22');
            set(cells.Font, 'Bold', true);
            set(cells.Interior,'ColorIndex',15);
        end
        
        if strcmp(db_production.raw.(currentcaseid).topmaterial,'EVA Carbon') || strcmp(db_production.raw.(currentcaseid).topmaterial,'EVA carbon')
            cells = Excel.ActiveSheet.Range('B24:D24');
            set(cells.Font, 'Bold', true);
            set(cells.Interior,'ColorIndex',15);
            %         if strcmp(db_production.raw.(currentcaseid).topthickness,'6 mm')
            %             cells = Excel.ActiveSheet.Range('B22:D22');
            %             set(cells.Font, 'Bold', true);
            %             set(cells.Interior,'ColorIndex',15);
            %         end
        end
        if strcmp(db_production.raw.(currentcaseid).topmaterial,'EVA Carbon') && strcmp(db_production.raw.(currentcaseid).company,'Livit Orthopedie bv')
            cells = Excel.ActiveSheet.Range('B27:D27');
            set(cells.Font, 'Bold', true);
            set(cells.Interior,'ColorIndex',15);
            set(cells, 'Value', 'No top layer');
        end
        if sum(strcmp(db_production.raw.(currentcaseid).heelpadleft,heelpads)) > 0
            cells = Excel.ActiveSheet.Range('B28:D28');
            set(cells.Font, 'Bold', true);
            set(cells.Interior,'ColorIndex',15);
        end
        if sum(strcmp(db_production.raw.(currentcaseid).heelpadright,heelpads)) > 0
            cells = Excel.ActiveSheet.Range('B29:D29');
            set(cells.Font, 'Bold', true);
            set(cells.Interior,'ColorIndex',15);
        end
        if strcmp(db_production.raw.(currentcaseid).servicetype,'Assembled - full length') == 0
            cells = Excel.ActiveSheet.Range('B27:D27');
            set(cells.Font, 'Bold', true);
            set(cells.Interior,'ColorIndex',15);
            if strcmp(db_production.raw.(currentcaseid).servicetype,'No top layer') == 1
                eActivesheetRange = get(Excel.Activesheet,'Range','B22:B26');
                eActivesheetRange.Value = '';
            end
        end
        
        % Customize the Excel file
        % ------------------------
        
        %     set(Excel.Selection.Font,'ColorIndex',7);
        %     set(Excel.Selection,'HorizontalAlignment',3);
        %     set(Excel.Selection.Interior,'ColorIndex',4);
        %     set(Excel.Selection.Font,'Size',13);
        %     set(Excel.Selection.Font,'bold',1);
        
        if strcmp(db_production.raw.(currentcaseid).casecancelled,'-') == 0
            cells = Excel.ActiveSheet.Range('A2:D5');
            cells.Select;
            cells.MergeCells = 1;
            cells.VerticalAlignment = 2;
            cells.HorizontalAlignment = 3;
            cells.WrapText = true;
            set(cells, 'Value', 'CANCELLED');
            cells.Font.Size = 45;
            
            cells = Excel.ActiveSheet.Range('F8:H40');
            cells.Select;
            cells.MergeCells = 1;
            cells.VerticalAlignment = 2;
            cells.HorizontalAlignment = 3;
            cells.WrapText = true;
            set(cells, 'Value', 'CANCELLED');
            cells.Font.Size = 75;
            cells.Orientation = -90;
        end
        
        cells = Excel.ActiveSheet.Range('B8');
        set(cells.Font, 'Bold', true);
        cells.Font.Size = 20;
        
        cells = Excel.ActiveSheet.Range('A42:D42');
        cells.Select;
        cells.MergeCells = 1;
        set(cells,'HorizontalAlignment',3);
        cells.Border.Item('xlEdgeTop').LineStyle = 1;
        
        cells = Excel.ActiveSheet.Range('A43:D43');
        cells.Select;
        cells.MergeCells = 1;
        set(cells,'HorizontalAlignment',3);
        
        cells = Excel.ActiveSheet.Range('A44:D44');
        cells.Select;
        cells.MergeCells = 1;
        set(cells,'HorizontalAlignment',3);
        
        cells = Excel.ActiveSheet.Range('A45:D45');
        cells.Select;
        cells.MergeCells = 1;
        set(cells,'HorizontalAlignment',3);
        
        cells = Excel.ActiveSheet.Range('A46:D46');
        cells.Select;
        cells.MergeCells = 1;
        set(cells,'HorizontalAlignment',3);
        cells.Border.Item('xlEdgeBottom').LineStyle = 1;
        
        cells = Excel.ActiveSheet.Range('B32:D40');
        cells.Select;
        cells.MergeCells = 1;
        cells.VerticalAlignment = 1;
        %set(cells, 'WrapText', 'True');
        cells.WrapText = true;
        
        cells = Excel.ActiveSheet.Range('A6:D6');
        cells.Border.Item('xlEdgeBottom').LineStyle = 1;
        cells = Excel.ActiveSheet.Range('A9:D9');
        cells.Border.Item('xlEdgeBottom').LineStyle = 1;
        cells = Excel.ActiveSheet.Range('A19:D19');
        cells.Border.Item('xlEdgeBottom').LineStyle = 1;
        cells = Excel.ActiveSheet.Range('A30:D30');
        cells.Border.Item('xlEdgeBottom').LineStyle = 1;
        
        % Set width of first column
        Sheet1.Columns.Item(2).HorizontalAlignment = 2;
        Sheet1.Columns.Item(1).columnWidth = 15;
        Sheet1.Columns.Item(2).columnWidth = 13;
        Sheet1.Columns.Item(3).columnWidth = 5.11;
        Sheet1.Rows.Item(7).rowHeight = 10;
        Sheet1.Rows.Item(9).rowHeight = 10;
        Sheet1.Rows.Item(10).rowHeight = 10;
        Sheet1.Rows.Item(19).rowHeight = 10;
        Sheet1.Rows.Item(20).rowHeight = 10;
        Sheet1.Rows.Item(30).rowHeight = 10;
        Sheet1.Rows.Item(31).rowHeight = 10;
        %    set('B17','HorizontalAlignment','xlHAlignRight');
        
        %     % Insert the QR code
        %     qr = qrcodegenerate(currentcaseid_long,'QuietZone',2);
        %     qr = encode_qr('Hello World','.*');
        %     fig = figure;
        %     colormap(gray);
        %     imagesc(qr);
        %     axis off;
        %     set(gcf, 'Position', [100, 100, 500, 440]);
        %     print(fig, ['Temp\' currentcaseid_long],'-dpng');
        %     img2 = ['Temp\' currentcaseid_long '.png'];
        %     Shapes.AddPicture([pwd '\' img2],0,1,1,450,80,80);
        %     close(fig);
        %
        %cells = Excel.ActiveSheet.Range('B17:B18');
        %displayFormat = cells.DisplayFormat;
        %cellstyle = displayFormat.Style;
        %set(cellstyle,'HorizontalAlignment','xlHAlignLeft');
        %Sheet1.Rows.Item(17).HorizontalAlignment = ;
        
        % Set borders
        %     range = invoke(Sheet1, 'Range', 'B10:G20');
        %     borders = get(range, 'Borders');
        %     set(borders, 'ColorIndex', 3);
        %     set(borders, 'LineStyle', 9);
        
        %     if strcmp(db_production.raw.(currentcaseid).delivery_country,'United States') == 1 ...
        %            || strcmp(db_production.raw.(currentcaseid).delivery_country,'Canada') == 1
        %
        %         destination_path = [pwd '\Output\FinishingReports\US\' currentcaseid '.xlsx'];
        %     else
        %         destination_path = [pwd '\Output\FinishingReports\BE\' currentcaseid '.xlsx'];
        %     end
        %     invoke(Workbook, 'SaveAs', destination_path);
        invoke(Workbook, 'SaveAs', [values.finrepfolder  currentcaseid '.xlsx']);
        % Save in the finrep folder
        %     if values.UpdateOnlineReport == 1
        %         copyfile(destination_path,[values.finrepfolder  currentcaseid '.xlsx']);
        %     end
        % % Print excel file
        % if values.MenuSelectPrinter > 1
        %     Excel.ActiveWorkbook.PrintOut(1,1,1,'False',values.MenuSelectPrinter_str);
        % end
        invoke(Excel, 'Quit');
    end
    
    % Create an overview of the cases and sort by priority
    overview(:,1) = db_production.overview(:,5);
    overview(:,2) = db_production.overview(:,7);
    overview(:,3) = db_production.overview(:,1);
    overview(:,4) = db_production.overview(:,12);
    overview(:,5) = db_production.overview(:,18);
    
    overview = sortrows(overview,[1,2,4]);
    
    xlswrite([pwd '\Output\FinishingReports\OverviewPrioritiesLastBatch.xlsx'],overview);
    
end

end