function [nextrow,db] = checkforfield(db,currentclient,fieldtofind)

% Check if the field already exists
if isfield(db.clientspecific.(currentclient),fieldtofind) == 1
    donothing = 1;
    nextrow = size(db.clientspecific.(currentclient).additionalproduct,1) + 1;
else
    db.clientspecific.(currentclient).(fieldtofind) = cell(1,12);
    nextrow = size(db.clientspecific.(currentclient).additionalproduct,1);
end

end