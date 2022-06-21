function [temp] = catchrowindex2(input,source,colnr)

% catchrowindex2 gives back ALL occurences

% input = list of strings you want to find
% source = source file with all the data
% colnr = target column where to find the strings

nrofstrings = size(input,2);
for cs = 1:nrofstrings
    string = char(input(1,cs));
    idx = strcmp(string,source(:,colnr));
    temp = find(idx==1);
    clear string
    clear idx
end

end