function [agrnr] = getcurrentpriceagreement(customers,cusnr,cy,cm,cd)

load Temp\values.mat values

% Get the number of agreements for this customer
if strcmp(cusnr,'C1181') == 1
    test = 1;
end
nrofagr = length(fieldnames(customers.(cusnr).PriceAgreements));

if nrofagr == 0
    disp(['No customer agreements found for ' cusnr ' - please check']);
else
    daterange = cell(nrofagr,9);
    for ca = 1:nrofagr
        daterange(ca,1) = {ca};
        daterange(ca,2) = {customers.(cusnr).PriceAgreements.(['Agreement' num2str(ca)]).StartYear};
        daterange(ca,3) = {customers.(cusnr).PriceAgreements.(['Agreement' num2str(ca)]).StartMonth};
        daterange(ca,4) = {customers.(cusnr).PriceAgreements.(['Agreement' num2str(ca)]).StartDay};
        daterange(ca,5) = {customers.(cusnr).PriceAgreements.(['Agreement' num2str(ca)]).EndYear};
        daterange(ca,6) = {customers.(cusnr).PriceAgreements.(['Agreement' num2str(ca)]).EndMonth};
        daterange(ca,7) = {customers.(cusnr).PriceAgreements.(['Agreement' num2str(ca)]).EndDay};
    end
    clear ca
    % Begin with arranging them in chronological order based on starting date.
    daterange = sortrows(daterange,[2,3,4]);
    % Now select the active agreement.
    for ca = 1:nrofagr
        % Check if the starting date is the same of before the creation date of the order
        if (cy == cell2mat(daterange(ca,2)) && cm == cell2mat(daterange(ca,3)) && cd >= cell2mat(daterange(ca,4))) || ...
           (cy == cell2mat(daterange(ca,2)) && cm > cell2mat(daterange(ca,3))) || ...
           (cy > cell2mat(daterange(ca,2)))
            % The current agreement is valid based on the starting date
            % Check if there is an end date, if yes, see if it hasn't passed yet.
            if isnumeric(cell2mat(daterange(ca,5))) == 1                
                if (cy >= cell2mat(daterange(ca,5)) && cm >= cell2mat(daterange(ca,6)) && cd > cell2mat(daterange(ca,7))) || ...
                   (cy >= cell2mat(daterange(ca,5)) && cm > cell2mat(daterange(ca,6))) || ...
                   (cy > cell2mat(daterange(ca,5)))
                    % The current agreement is not valid based on the ending date
                    daterange(ca,8) = {'Too late to use this agreement'};
                    daterange(ca,9) = {0};
                else
                    % The current agreement is also valid based on the ending date
                    if ca < nrofagr
                        if (cy == cell2mat(daterange(ca+1,2)) && cm == cell2mat(daterange(ca+1,3)) && cd >= cell2mat(daterange(ca+1,4))) || ...
                                (cy == cell2mat(daterange(ca+1,2)) && cm > cell2mat(daterange(ca+1,3))) || ...
                                (cy > cell2mat(daterange(ca+1,2)))
                            daterange(ca,8) = {'Overruled by next agreement'};
                            daterange(ca,9) = {0};
                        else
                            daterange(ca,8) = {'Possible 1'};
                            daterange(ca,9) = {1};
                        end
                    else
                        % If there is no next agreement, it can't be overruled.
                        daterange(ca,8) = {'Possible 2'};
                        daterange(ca,9) = {1};
                    end
                end
            else
                % Check if the next agreement will overrule the current one
                if ca < nrofagr
                    if (cy == cell2mat(daterange(ca+1,2)) && cm == cell2mat(daterange(ca+1,3)) && cd >= cell2mat(daterange(ca+1,4))) || ...
                       (cy == cell2mat(daterange(ca+1,2)) && cm > cell2mat(daterange(ca+1,3))) || ...
                       (cy > cell2mat(daterange(ca+1,2)))                        
                        daterange(ca,8) = {'Overruled by next agreement'};    
                        daterange(ca,9) = {0};
                    else
                        daterange(ca,8) = {'Possible 3'};
                        daterange(ca,9) = {1};
                    end
                else
                    % If there is no next agreement, it can't be overruled.
                    daterange(ca,8) = {'Possible 4'};
                    daterange(ca,9) = {1};
                end
            end
        else
            daterange(ca,8) = {'Too soon to use this agreement'};
            daterange(ca,9) = {0};
        end
    end
    % Search for the last possible agreement. Normally there should be only (max) one 'Possible' match.    
    matches = cell2mat(daterange(:,9));
    if sum(matches) == 0
        disp(['No agreement is possible for ' cusnr ' . Add the correct agreement in the Excel or check the agreements to see what is going on.']);
        agrnr = 0;
    elseif sum(matches) == 1
        temp = find(matches==1);
        % Get the correct agreement number.
        agrnr = cell2mat(daterange(temp,1));
        % disp(['Agreement ' num2str(agrnr) ' is the correct one for customer ' cusnr '.']);
    elseif sum(matches) > 1
        disp(['More than one agreement is possible for ' cusnr '. Check the agreements to see what is going on.']);
        clear all
    else
        disp(['There seems to be an issue with finding the correct agreement for ' cusnr '.']);
        clear all
    end
    
end

end