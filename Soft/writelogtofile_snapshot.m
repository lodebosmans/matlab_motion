function writelogtofile_snapshot(caseids_production,timestring,testdebug,values,procedure)

% Write the daily production snapshot to file to file

% See if a file already exist for this case ID or RS-number
if length(testdebug) > 1
    filename = [values.logfolder 'DailyProductionSnapshot\' values.y values.mo values.d '_' values.h values.mi values.s '_DailyProductionSnapshot_TEST.txt'];
else
    filename = [values.logfolder 'DailyProductionSnapshot\' values.y values.mo values.d '_' values.h values.mi values.s '_DailyProductionSnapshot.txt'];
end

if isfile(filename)
    % File exists. Open it.
    fid = fopen(filename,'at');
else
    % File does not exist. Create it.
    fid = fopen(filename,'wt');
    fprintf(fid, [timestring testdebug 'Creation of daily production snapshot file [' values.Username_str ']']);
    fprintf(fid, '\n');
    fprintf(fid, '\n');
end
% Write the action to file.
fprintf(fid, [procedure ' (' num2str(size(caseids_production,1)) ')']);
fprintf(fid, '\n');
for x = 1:size(caseids_production,1)
    fprintf(fid, char(caseids_production(x,1)));
    fprintf(fid, '\n');
end
fprintf(fid, '\n');
fclose(fid);

end