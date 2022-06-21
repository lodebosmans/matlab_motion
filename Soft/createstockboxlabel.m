function createstockboxlabel()

load Temp\values.mat values

% Read the excel file with the data
[NUM,TXT,RAW] = xlsread(values.stockboxlabelfolder);

nrofitems = size(RAW,1);
% Create the result variable
labels = '';

% Go over all items listed. Skip the header.
for x = 2:nrofitems
    % Set the default value for the pairs-label
    pairlabel = 'Pairs';
    % Get the itemnumber
    itemnumber = char(RAW(x,1));
    % Get the amount
    amount = cell2mat(RAW(x,2));
    % Get the size of the amount, if larger than 999, downsize the QR-code size
    if amount > 999
        % Smaller size
        QRscale = '8';
    else
        % Normal size
        QRscale = '10';
    end
    % Make amount a string.
    amount = num2str(cell2mat(RAW(x,2)));
    
    % Check what material type is included
    if strcmp(itemnumber(1),'1') == 1
        material = 'EVA';
    elseif strcmp(itemnumber(1),'2') == 1
        material = 'PU soft';
    elseif strcmp(itemnumber(1),'3') == 1
        material = 'EVA carbon';
    else
        error('Top layer material not known');
    end
    
    % Check what type of top layer material
    if strcmp(itemnumber(2),'0') == 1
        % Regular type
        if strcmp(itemnumber(1),'2') == 1
            type = 'Black fabric';
        else
            type = '-';
        end
    elseif strcmp(itemnumber(2),'1') == 1
        % Special type
        type = 'Synth. leather';
    else
        error('Top layer type not known');
    end
    
    % Get the thickness
    thickness = itemnumber(3);
    
    % Get the size
    if strcmp(itemnumber(12),'S') == 1
        % Overrule if it is a sheet
        sizing = 'sheet';
        pairlabel = 'Sheets';
    else
        if strcmp(itemnumber(7),'0') == 1
            % For a single character size
            sizing = itemnumber(8);
        else
            % For a double character size
            sizing = itemnumber(7:8);
        end
    end
    
    % Get the shore value
    shore = itemnumber(10:11);
    
    % Get the date
    if values.d < 10
        day = ['0' num2str(values.d)];
    else
        day = num2str(values.d);
    end
    
    if values.mo < 10
        month = ['0' num2str(values.mo)];
    else
        month = num2str(values.mo);
    end
    
    year = num2str(values.y);
    
    
     
    
    % Label template and insert the 
    label_temp = ['^XA' ...
        '^FX Main grid' ...
        '^FO50,50^GB700,1100,3^FS' ...
        '^FX Grid for company name and date' ...
        '^FO50,50^GB700,150,3^FS' ...
        '^FO50,50^GB350,150,3^FS' ...
        '^FX Values' ...
        '^CF0,60' ...
        '^FO75,100^FDRS Print HQ^FS' ...
        '^FO435,100^FD' day '/' month '/' year '^FS' ...
        '^FX Grid for material property name' ...
        '^FO50,50^GB700,650,3^FS' ...
        '^FX Values' ...
        '^CF0,45' ...
        '^FO75,230^FDMaterial^FS' ...
        '^FO75,330^FDType^FS' ...
        '^FO75,430^FDThickness (mm)^FS' ...
        '^FO75,530^FDSize (UK)^FS' ...
        '^FO75,630^FDShore^FS' ...
        '^FX Grid for material property value' ...
        '^FO50,50^GB350,650,3^FS' ...
        '^FX Values' ...
        '^FO435,230^FD' material '^FS' ...
        '^FO435,330^FD' type '^FS' ...
        '^FO435,430^FD' thickness '^FS' ...
        '^FO435,530^FD' sizing '^FS' ...
        '^FO435,630^FD' shore '^FS' ...
        '^FX Grid for itemnumber' ...
        '^FO50,50^GB700,850,3^FS' ...
        '^FO75,720^FDItemnumber^FS' ...
        '^CF0,70' ...
        '^FO200,790^FD' itemnumber '^FS' ...
        '^FX Values pairs' ...
        '^CF0,45' ...
        '^FO75,920^FD' pairlabel '^FS' ...
        '^CF0,100' ...
        '^FO150,1000^FD' amount '^FS' ...
        '^FO450,910^BQN,2,' QRscale '^FDQA,' itemnumber ',' amount '^FS' ...
        '^XZ'];
    
    labels = [labels label_temp];
end

labels

% Create text file for printing label code
fid = fopen('Output\Labels\results_labels.txt','wt');
fprintf(fid, '${');
fprintf(fid, '\n');
fprintf(fid, labels);
fprintf(fid, '}$');
fclose(fid);

try
    web('https://supportcommunity.zebra.com/s/article/Sending-ZPL-Commands-to-a-Printer?language=en_US');
catch
end



end