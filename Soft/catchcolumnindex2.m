function [input] = catchcolumnindex2(input,source,rownr)

% catchcolumnindex2 gives back ALL occurences

% input = list of strings you want to find (columnvector)
% source = source file with all the data
% rownr = target row where to find the strings

nrofstrings = size(input,2);
for cs = 1:nrofstrings
    string = char(input(1,cs));
    idx = strcmp(string,source(rownr,:));
    input(2,cs) = num2cell(find(idx==1));
    clear string
    clear idx
end

end