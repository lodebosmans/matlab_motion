function [input] = catchrowindexall(input,source,colnr)

% input = list of strings you want to find
% source = source file with all the data
% colnr = target column where to find the strings

nrofstrings = size(input,2);
for cs = 1:nrofstrings
    string = char(input(1,cs));
    idx = strcmp(string,source(:,colnr));
    temp = num2cell(find(idx==1));
    input = cell2mat(temp);
    clear string
    clear idx
end

end