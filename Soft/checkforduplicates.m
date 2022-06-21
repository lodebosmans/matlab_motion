function checkforduplicates(caseids_production)

load Temp\values.mat values

strings = caseids_production';
nrofcaseIDs = size(strings,2);
[~, uniqueIdx] = unique(strings); % Find the indices of the unique strings
duplicates = strings; % Copy the original into a duplicate array
duplicates(uniqueIdx) = []; % remove the unique strings, anything left is a duplicate
duplicates = unique(duplicates); % find the unique duplicates
nrofduplicates = size(duplicates,2);
if nrofduplicates > 0
    if nrofduplicates == 1
        % Sentence for one duplicate.
        message = ['There is ' num2str(nrofduplicates) ' duplicate case ID in the total of ' num2str(nrofcaseIDs) ' provided case IDs. There are thus only ' num2str(nrofcaseIDs - nrofduplicates) ' unique case IDs instead of ' num2str(nrofcaseIDs) '.' newline newline 'Please verify and retry with the correct amount of case IDs.' newline newline 'The duplicate case ID is: '];
    else
        % Sentence for multiple duplicates.
        message = ['There are ' num2str(nrofduplicates) ' duplicate case IDs in the total of ' num2str(nrofcaseIDs) ' provided case IDs. There are thus only ' num2str(nrofcaseIDs - nrofduplicates) ' unique case IDs instead of ' num2str(nrofcaseIDs) '.' newline newline 'Please verify and retry with the correct amount of case IDs.' newline newline 'The duplicate case IDs are: '];    
    end
    for cdup = 1:nrofduplicates
        ccid = char(duplicates(1,cdup));
        if cdup < nrofduplicates
            message = [message ccid ', ' ];
        else
            message = [message ccid '.' ];
        end 
    end   
    % Show error message
    errordlg(message);  
    error('Too many duplicate case IDs...')
end

end