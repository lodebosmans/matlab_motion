function [filename] = checkiffilenameexists(filename,extension)

goon = 0;

while goon == 0
    
    if isfile([filename extension]) == 1
        filename = [filename '_1'];
    else
        goon = 1;
    end
    
end
    
end