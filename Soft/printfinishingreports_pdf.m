function printfinishingreports_pdf()

% https://nl.mathworks.com/matlabcentral/answers/103861-how-can-i-save-a-word-document-as-a-pdf-via-actxserver-in-matlab-8-2-r2013b

load Temp\values.mat values
% load Temp\db_production.mat db_production
load Temp\caseids_production.mat caseids_production

% Go over all finishing reports and print them
if values.MenuSelectPrinter > 1
    for x = 1:size(caseids_production,1)
        
        ccid = strrep(char(caseids_production(x,1)),'-','');
        
        filename = [values.finrepfolder  ccid '.docx']; % This must be full path name  => only adapt the file in the 'Interface' folder, not in the test environment
        
%         word = actxserver('Word.Application');      %start Word
%         word.Visible = 1;                           %make Word Visible
%         document = word.Documents.Open(filename);
%         
%         % Send to printer
%         
%         
%         word.Run('add_my_header', sprintf('%s %s', datestr(now), filename)) 
%         
%         document.PrintOut(1,1,1,filename,values.MenuSelectPrinter_str);
%         %         Excel.ActiveWorkbook.PrintOut(1,1,1,'False',values.MenuSelectPrinter_str);
%         
%         % Close Excel and clean up
% %         invoke(Excel,'Quit');
% %         delete(Excel);
% %         clear Excel;

        filepath = 'C:\Users\mathlab\Documents\Matlab\20190318_Interface\Input\Templates\FinRep.docx';
        pythonfile = 'C:\Users\mathlab\Documents\Matlab\20190318_Interface_OrFix\Soft\winprint.py';

        command = ['python ' pythonfile ' ' filepath];
        
        system(command);



    end
end

end