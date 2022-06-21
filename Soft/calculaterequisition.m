function calculaterequisition()

load Temp\values.mat values

% Load the file with the requisition data
files = dir('Input\*Requis*.xlsx');
directory.requis = [files(1).folder '\' files(1).name];
[num,text,requisinput] = xlsread(directory.requis,1);
nrofrows = size(requisinput,1);

% Get the columns of the crucial data
col_itemnumber = catchcolumnindex({'No.'},requisinput,3);
col_itemnumber = cell2mat(col_itemnumber(2,1));
col_location = catchcolumnindex({'Location Code'},requisinput,3);
col_location = cell2mat(col_location(2,1));
col_quantity = catchcolumnindex({'Quantity'},requisinput,3);
col_quantity = cell2mat(col_quantity(2,1));

% Scan all locations

% Create a struct with all the locations to write the transitions
itemjournal.BEPROD = struct;
itemjournal.USPROD = struct;
counter = 0;

for cr = 4:nrofrows
    cloc = char(requisinput(cr,col_location));
    cin = char(requisinput(cr,col_itemnumber));
    cq = cell2mat(requisinput(cr,col_quantity));
    
    if length(cin) == 12
        material = cin(1:2);
        height = cin(3:4);
        shore = cin(10:11);
        
        if strcmp(cin(1:2),'10') == 1
            eva = 1
        elseif strcmp(cin(1:2),'11') == 1
            
        elseif strcmp(cin(1:2),'20') == 1
            
        elseif strcmp(cin(1:2),'21') == 1
            
        elseif strcmp(cin(1:2),'21') == 1
            
        elseif strcmp(cin(1:2),'30') == 1
            
        end
    end
    
    
    
end


end