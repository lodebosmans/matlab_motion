function writelogtofile(ccid,timestring,phase,testdebug,values,procedure)

% Write the logevent to file
disp(['Writing event in logfile for ' ccid ' - please wait']);

% See if a file already exist for this case ID or RS-number
filename = [values.logfolder 'Cases\' ccid '.txt'];

if isfile(filename)
    % File exists. Open it.
    fid = fopen(filename,'at');
else
    % File does not exist. Create it.
    fid = fopen(filename,'wt');
    fprintf(fid, [timestring testdebug 'Creation of logfile ['  values.Username_str ']']);
    fprintf(fid, '\n');
end
% Write the action to file.
fprintf(fid, [timestring testdebug procedure  ' ['  phase ' ' values.Username_str ']']);
fprintf(fid, '\n');
fclose(fid);

end