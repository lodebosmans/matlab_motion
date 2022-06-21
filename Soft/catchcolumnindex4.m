function [input] = catchcolumnindex4(input,source,rownr,mode)

% input = list of strings you want to find (in cell format)  (columnvector)
% source = source file with all the data
% rownr = target row where to find the input in the source
% mode = indicates the mode: find the (1) first, (2) last or (3) all occurences

nrofstrings = size(input,2);
for cs = 1:nrofstrings
    string = char(input(1,cs));
    idx = strcmp(string,source(rownr,:));
    input(2,cs) = num2cell(find(idx==1));
    clear string
    clear idx
end

end