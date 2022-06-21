function [db] = reducedata(db,source,destination,endcolumn,stringstofetch_invoice)

for y = 1:endcolumn
    eval(['db.' destination '(:,y) = db.' source '(:,cell2mat(stringstofetch_invoice(2,y)));']);
end
clear y

% Replace the top layer with only the info of the additional synthetic top layer (and the additional cost related to it)
eval([ 'db.' destination '(cellfun(@(x) strcmp(x,''EVA Synthetic Leather'' ), db.' destination ')) = {''Synthetic leather''};']);
eval([ 'db.' destination '(cellfun(@(x) strcmp(x,''PU Soft Synthetic Leather''), db.' destination ')) = {''Synthetic leather''};']);
eval([ 'db.' destination '(cellfun(@(x) strcmp(x,''EVA''), db.' destination ')) = {''''};']);
eval([ 'db.' destination '(cellfun(@(x) strcmp(x,''PU Soft''), db.' destination ')) = {''''};']);

% Filter out the NaNs
eval([ 'db.' destination '(cellfun(@(x) isnumeric(x) && isnan(x), db.' destination ')) = cellstr('''');']);

end