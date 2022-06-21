function [values] = getdailysnapshot(values)

% Some default parameters
%values.snapshotprocedures = {'prod_received','prod_coloring','prod_inprocess','prod_readytoship'};
values.snapshotprocedures_text = {'Insoles received','Coloring/drying','In finishing process','Ready to ship'};
values.snapshotprocedures = {'Received','Coloring','FinishingProcess','ReadyToShip'};
values.snapshotprocedures2 = {'Received','Coloring','FinishingProcess','ReadyToShip','RecColCombo'};
inputstrings = {'Insert case IDs here.. '};
% Gather all the data
goon = 0;
while goon == 0
    prompt = 'Enter the statusses and case IDs here.';
    title = 'Daily production status snapshot';
    dims = [20 100];
    definput = inputstrings;
    values.snapshotallcaseids = char(inputdlg(prompt,title,dims,definput));
    % Check for duplicates, if so return the input
    [caseids_production] = readmanualcaseids({},'',0,values.snapshotallcaseids,0);
    [goon] = checkforduplicates2return(caseids_production);
    if goon == 0
        inputstrings = {values.snapshotallcaseids};
    end
end
% Define all statusses
statusoccurences = {};
% Start loop here
for cprod = 1:size(values.snapshotprocedures2,2)
    cprodstr = char(values.snapshotprocedures2(1,cprod));
    values.(cprodstr) = 'RS19-PHI-ITS';
    indexes = strfind(values.snapshotallcaseids,cprodstr);
    nrofoccurences = size(indexes,2);
    if nrofoccurences > 0
        for co = 1:nrofoccurences
            if strcmp(cprodstr,'RecColCombo') == 0
                occrow = size(statusoccurences,1);
                statusoccurences(occrow+1,1) = {indexes(1,co)}; %#ok<*AGROW>
                statusoccurences(occrow+1,2) = {cprodstr};
            else
                occrow = size(statusoccurences,1);
                statusoccurences(occrow+1,1) = {indexes(1,co)};
                statusoccurences(occrow+1,2) = {'Received'};
                statusoccurences(occrow+2,1) = {indexes(1,co)};
                statusoccurences(occrow+2,2) = {'Coloring'};
            end
        end
    end
end
% Now sort them per procedure
statusoccurences = sortrows(statusoccurences,1);
stringlength = length(values.snapshotallcaseids);
nrofrows = size(statusoccurences,1);
for x = 1:nrofrows
    disp(['Row ' num2str(x)]);
    cprod = char(statusoccurences(x,2));
    if x < nrofrows
        startpos = cell2mat(statusoccurences(x,1));
        endpos = cell2mat(statusoccurences(x+1,1));
        if startpos == endpos
            counter = 1;
            while startpos == endpos    
                if x + counter <= nrofrows
                    endpos = cell2mat(statusoccurences(x + counter,1));
                else
                    endpos = stringlength;
                end
                counter = counter + 1;
            end
        end
    else
        startpos = cell2mat(statusoccurences(x,1));
        endpos = stringlength;
    end
    % Get the case IDs
    temp = readmanualcaseids({},'',0,values.snapshotallcaseids(startpos:endpos),0);
    nrofcaseids = size(temp,1);
    if nrofcaseids > 0
        for y = 1:nrofcaseids
            values.(cprod) = [values.(cprod) ', '  char(temp(y,1))];
        end
    end    
end


end