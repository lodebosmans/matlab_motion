function createfinishingreports_pdf()

load Temp\caseids_production_tocreate.mat caseids_production_tocreate

if size(caseids_production_tocreate,1) > 0
    
    load Temp\values.mat values
    load Temp\db_production.mat db_production
    
    
    nrofcases = size(db_production.overview,1);
    heelpads = {'Plantar Fasciitis','Heel Spur','Fat Pad'};
    filename = [values.deliverynotetemplatefolder 'FinRep.docx'];
    
    for caseindex = 1:nrofcases
        
        % Get the caseID from the list of caseIDs
        clear currentcaseid
        currentcaseid_long = char(db_production.overview(caseindex,1));
        currentcaseid = [currentcaseid_long(1:4) currentcaseid_long(6:8) currentcaseid_long(10:12)];
        
        displaytext = ['Creating finishing reports - please wait - processing ' num2str(caseindex) ' of ' num2str(nrofcases) ' - ' currentcaseid_long];
        disp(displaytext);
        
        word = actxserver('Word.Application');      %start Word
        word.Visible = 0;                           %make Word Visible
        %         document=word.Documents.Add;                %create new Document
        document = word.Documents.Open(filename);
        selection=word.Selection;                   %set Cursor
        
        %         selection.Font.Name='Courier New';          %set Font
        %         selection.Font.Size=9;                      %set Size
        %         margin = 28.34646;
        %         selection.Pagesetup.RightMargin = margin;   %set right Margin to 1cm
        %         selection.Pagesetup.LeftMargin = margin;    %set left Margin to 1cm
        %         selection.Pagesetup.TopMargin = margin;     %set top Margin to 1cm
        %         selection.Pagesetup.BottomMargin = margin;  %set bottom Margin to 1cm
        %         selection.Paragraphs.LineUnitAfter=0.00;    %sets the amount of spacing between paragraphs(in gridlines)
        %         selection.InlineShapes.AddPicture([values.deliverynotetemplatefolder 'Materialise_BL_sRGB.png'],0,1);


        
        gray_index = 16;
        notoplayer = strcmp(db_production.raw.(currentcaseid).servicetype,'No top layer');
        

        selection.MoveDown(5,1);
        selection.MoveDown(5,1);
        selection.Font.Size = 20;
        selection.TypeText(db_production.raw.(currentcaseid).caseid.full);
        selection.MoveDown(5,1);
        selection.MoveDown(5,1);
        selection.TypeText(db_production.raw.(currentcaseid).surgeon);
        selection.MoveDown(5,1);
        selection.TypeText(db_production.raw.(currentcaseid).delivery_company);
        selection.MoveDown(5,1);
        selection.TypeText(db_production.raw.(currentcaseid).delivery_street);
        selection.MoveDown(5,1);
        if strcmp(db_production.raw.(currentcaseid).delivery_state,'<None>')
            selection.TypeText(db_production.raw.(currentcaseid).delivery_zip_city);
        else
            selection.TypeText([db_production.raw.(currentcaseid).delivery_zip_city ', ' db_production.raw.(currentcaseid).delivery_state]);
        end
        selection.MoveDown(5,1);
        selection.TypeText(db_production.raw.(currentcaseid).delivery_country);
        selection.MoveDown(5,1);
        selection.TypeText(db_production.raw.(currentcaseid).estimatedshippingdate);
        selection.MoveDown(5,1);
        selection.TypeText(db_production.raw.(currentcaseid).referenceid);
        selection.MoveDown(5,1);
        selection.MoveDown(5,1);
        selection.TypeText(db_production.raw.(currentcaseid).insoletype);
        selection.MoveDown(5,1);
        if notoplayer == 0
            if strcmp(db_production.raw.(currentcaseid).heelcupleft,'Low') || strcmp(db_production.raw.(currentcaseid).heelcupright,'Low')
                selection.Font.Bold = 1;
                selection.Shading.BackgroundPatternColorindex= gray_index;
            end
            selection.TypeText([db_production.raw.(currentcaseid).heelcupleft ' - ' db_production.raw.(currentcaseid).heelcupright]);
            selection.MoveDown(5,1);
            selection.TypeText(db_production.raw.(currentcaseid).topthickness);
            selection.MoveDown(5,1);
            if strcmp(db_production.raw.(currentcaseid).topmaterial,'EVA Carbon') || strcmp(db_production.raw.(currentcaseid).topmaterial,'EVA carbon')
                selection.Font.Bold = 1;
                selection.Shading.BackgroundPatternColorindex = gray_index;
            end
            selection.TypeText(db_production.raw.(currentcaseid).topmaterial);
            selection.MoveDown(5,1);
            selection.TypeText(db_production.raw.(currentcaseid).topsize);
            selection.MoveDown(5,1);
            selection.TypeText(db_production.raw.(currentcaseid).tophardness);
            selection.MoveDown(5,1);
        else
            selection.MoveDown(5,1);
            selection.MoveDown(5,1);
            selection.MoveDown(5,1);
            selection.MoveDown(5,1);
            selection.MoveDown(5,1);
        end
        if strcmp(db_production.raw.(currentcaseid).servicetype,'Assembled - full length') == 0
            selection.Font.Bold = 1;
            selection.Shading.BackgroundPatternColorindex = gray_index;
        end
        selection.TypeText(db_production.raw.(currentcaseid).servicetype);
        selection.MoveDown(5,1);
        if notoplayer == 0
            if sum(strcmp(db_production.raw.(currentcaseid).heelpadleft,heelpads)) > 0
                selection.Font.Bold = 1;
                selection.Shading.BackgroundPatternColorindex = gray_index;
            end
            selection.TypeText(db_production.raw.(currentcaseid).heelpadleft);
            selection.MoveDown(5,1);
            if sum(strcmp(db_production.raw.(currentcaseid).heelpadright,heelpads)) > 0
                selection.Font.Bold = 1;
                selection.Shading.BackgroundPatternColorindex = gray_index;
            end
            selection.TypeText(db_production.raw.(currentcaseid).heelpadright);
        end
        
        destination_filename = [values.finrepfolder  currentcaseid];
        invoke(document, 'SaveAs', [destination_filename '.docx']);
        % invoke(document, 'SaveAs2', [destination_filename '.pdf'],17); % https://nl.mathworks.com/matlabcentral/answers/103861-how-can-i-save-a-word-document-as-a-pdf-via-actxserver-in-matlab-8-2-r2013b
        invoke(word, 'Quit');
        
    end
    
end