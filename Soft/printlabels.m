function printlabels()

load Temp\values.mat values
load Temp\caseids_production.mat caseids_production

%nrofcases = size(caseids_production,1);
nrofcases = size(caseids_production,1);

% Create text file for printing label code
fid = fopen('Output\Labels\results_labels.txt','wt');
fprintf(fid, '${');
fprintf(fid, '\n');

% Create for loop to iterate over all case IDs
for caseindex = 1:nrofcases
    displaytext = ['Combining labels - please wait - processing ' num2str(caseindex) ' of ' num2str(nrofcases)];
    disp(displaytext);
    % Get the caseID from the list of caseIDs
    clear currentcaseid
    currentcaseid_long = char(caseids_production(caseindex,1));
    currentcaseid = strrep(currentcaseid_long,'-','');
    
    textlabel = fileread([values.finrepfolder  currentcaseid '.txt']);
    % Remove the Zebra passthrough signs
    textlabel = strrep(textlabel,'${','');
    textlabel = strrep(textlabel,'}$','');
    
    if caseindex < nrofcases
        % Do not cut
        %         fprintf(fid, '^FS^XB^XZ');
        % Add the hold-cutting-signal
        textlabel = strrep(textlabel,'^FS^XZ','^FS^XB^XZ');
    end
    if caseindex == nrofcases
        % Cut
        %         fprintf(fid, '^FS^XZ');
        % Do not add the hold-cutting-signal
    end
    fprintf(fid, textlabel);
    fprintf(fid, '\n');
    fprintf(fid, '\n');
    
    clear currentcaseid
    clear code
end
fprintf(fid, '\n');
fprintf(fid, '}$');
fclose(fid);

end