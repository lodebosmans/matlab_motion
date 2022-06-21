function [output] = readmanualcaseids(output,fieldname,readfromvalues,content,ignorers)

% output = cell where case IDs should end up in. If no startcell is present, provide an empty cell {}.
% fieldname = fieldname of the values field the case ID's are in.
% readfromvalues = 1: read from values; 0: read from content.
% content = alternative to values, cell where case IDs are located in.
% ignorers =
%      0: Read case IDs & read RS numbers.
%      1: Read case IDs & ignore RS-numbers.
%      2: Do not read case IDs & only read RS-numbers.

if readfromvalues == 1
    load Temp\values.mat values
    % Convert all manually inserted case IDs to upper case.
    content = upper(values.(fieldname));
elseif readfromvalues == 0
    content = upper(content);
end

if ignorers == 0 || ignorers == 1
    prefixes = ["RS1","RS2"];
    for cpf_idx = 1:size(prefixes,2)
        cprefix = char(prefixes(1,cpf_idx));
        % Get the indexes of the start of the manually inserted case IDs
        indexes = strfind(content,cprefix);
        nrofmanualcases = size(indexes,2);
        for currentcase = 1:nrofmanualcases
            % Get current size of the list of case IDs
            currentsize = size(output,1);
            % Get the case ID
            startindex = indexes(1,currentcase);
            newcase = content(1,startindex:startindex+11);
            % Add it to the list
            output(currentsize+1,1) = cellstr(newcase);
        end
    end
    clear prefixes
    clear cprefix
    clear cpf_idx
    
%     % Get the indexes of the start of the manually inserted case IDs
%     indexes = strfind(content,'RS2');
%     nrofmanualcases = size(indexes,2);
%     for currentcase = 1:nrofmanualcases
%         % Get current size of the list of case IDs
%         currentsize = size(output,1);
%         % Get the case ID
%         startindex = indexes(1,currentcase);
%         newcase = content(1,startindex:startindex+11);
%         % Add it to the list
%         output(currentsize+1,1) = cellstr(newcase);
%     end
end

rscases.nr = 0;
if ignorers == 0 || ignorers == 2
    prefixes = ["RS-1","RS-2","FS-2"];
    for cpf_idx = 1:size(prefixes,2)
        cprefix = char(prefixes(1,cpf_idx));
        % Do the same for the numbers for marketing, retour, ...
        indexes = strfind(content,cprefix);
        nrofmanualcases = size(indexes,2);
        for currentcase = 1:nrofmanualcases
            % Get current size of the list of case IDs
            currentsize = size(output,1);
            % Get the RS/FS number
            startindex = indexes(1,currentcase);
            newcase = content(1,startindex:startindex+12);
            % Add it to the list
            output(currentsize+1,1) = cellstr(newcase);
        end
        rscases.nr = rscases.nr + nrofmanualcases;
    end
    
%     % Do the same for the numbers for marketing, retour, ...
%     indexes = strfind(content,'RS-2');
%     nrofmanualcases = size(indexes,2);
%     for currentcase = 1:nrofmanualcases
%         % Get current size of the list of case IDs
%         currentsize = size(output,1);
%         % Get the RT number
%         startindex = indexes(1,currentcase);
%         newcase = content(1,startindex:startindex+12);
%         % Add it to the list
%         output(currentsize+1,1) = cellstr(newcase);        
%     end
%     rscases.nr = rscases.nr + nrofmanualcases;
end

if rscases.nr > 0
    rscases.ids = output(end-rscases.nr+1:end,1);
end
save Temp\rscases.mat rscases

end