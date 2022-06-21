function [output] = readmanualcaseids2(values,input)

%Create a new output variable
output = cell(0,1);
% Convert all manually inserted case IDs to upper case.
input = upper(input);
% Get the indexes of the start of the manually inserted case IDs
indexes = strfind(input,'RS');
nrofmanualcases = size(indexes,2);
for currentcase = 1:nrofmanualcases
    % Get current size of the list of case IDs
    currentsize = size(output,1);
    % Get the case ID
    startindex = indexes(1,currentcase);
    newcase = input(1,startindex:startindex+11);
    % Add it to the list
    output(currentsize+1,1) = cellstr(newcase);
end

end