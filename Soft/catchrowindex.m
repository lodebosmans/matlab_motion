function [input] = catchrowindex(input,source,colnr)

% catchrowindex ONLY gives back THE LAST occurence

% input = list of strings you want to find
% source = source file with all the data
% colnr = target column where to find the strings

nrofstrings = size(input,2);
for cs = 1:nrofstrings
    string = char(input(1,cs));
    idx = strcmp(string,source(:,colnr));
    temp = num2cell(find(idx==1));
    input(2,cs) = temp(end,1);
    clear string
    clear idx
end

end