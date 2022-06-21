

system('"C:\Program Files\Microsoft Office\Office15\Winword.exe" /mFilePrintDefault C:\ExampleFile.docx &')


print /D:\\Network\Printer file

print /D:\\RSPRINTDC.RSPRINT.local\zebragrootetiket "C:\Users\mathlab\Documents\MATLAB\Interface_TestEnvironment\Output\results_labels.txt"

doc actxcontrolselect

print("/d:\\RSPRINTDC\zebragrootetiket","C:\Users\mathlab\Documents\MATLAB\Interface_TestEnvironment\Output\results_labels.txt")

print("C:\Users\mathlab\Documents\MATLAB\Interface_TestEnvironment\Output\results_labels.txt",'-Pzebragrootetiket')
%     PRINT(printer, ...) prints the figure or model to the specified printer.
%     Use '-P' to specify the printer option.
%       print(fig, '-Pmyprinter'); % print to the printer named 'myprinter'


    % Get handle to Excel COM Server
    NotepadPP = actxserver('NotepadPlusPlus.Application');
    NotepadPP.editors(1).activate()
    NotepadPP.activeEditor.openFile("C:\Users\mathlab\Documents\MATLAB\Interface_TestEnvironment\Output\results_labels.txt")
    NotepadPP.activeEditor.closeFile()
    
    "C:\Program Files (x86)\Notepad++\Notepad++.exe" /t "C:\Users\mathlab\Documents\MATLAB\Interface_TestEnvironment\Output\results_labels.txt" "Zebragrootetiket"
    "C:\Program Files (x86)\Notepad++\Notepad++.exe" "C:\Users\mathlab\Documents\MATLAB\Interface_TestEnvironment\Output\results_labels.txt"
    
    
    oNppApplication.activeEditor.openFile("C:\boot.ini")
    % Add a Workbook
    Workbooks = Excel.Workbooks;
    Workbook = invoke(Workbooks, 'Add');
    % Get a handle to Sheets and select Sheet 1
    Sheets = Excel.ActiveWorkBook.Sheets;
    Sheet1 = get(Sheets, 'Item', 1);
    Sheet1.Activate;