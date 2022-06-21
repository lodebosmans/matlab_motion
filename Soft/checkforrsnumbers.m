function checkforrsnumbers()

load Temp\values.mat values
load Temp\caseids_production.mat caseids_production
load Temp\rscases.mat rscases

clc
disp('Checking if all RS-number files are present. Please wait.')

rscases.missingnr = 0;
rscases.missingids = cell(0);

if rscases.nr > 0
    for crsc = 1:rscases.nr
        % Get the case ID
        ccid = char(rscases.ids(crsc,1));   
        % Check if the file is there
        if isfile([values.rsfilesfolder ccid '.xlsx']) == 1
            % File exists. Do nothing.
        else
            % File does not exist. Add to the list of missing RS-files.
            rscases.missingnr = rscases.missingnr  + 1;
            rscases.missingids(rscases.missingnr,1) = {ccid};
            disp(['The Excel file for ' ccid ' is missing. Please provide file.']);
        end
    end 
end
% Now give an error message if a RS-number is missing in the folder
if rscases.missingnr > 0
    save Temp\rscases.mat rscases
    errordlg('Sorry, not all RS-number files are available. See command window for the list of files that still need to be provided.');  
    error('See RS-numbers above.. these still need to be provided before being able to do the shipments.');
end

end