
filename = 'O:\Administration\BackupFiles\LogFiles\Cases\RS19-ACU-QAC.txt';


fileID = fopen('O:\Administration\BackupFiles\LogFiles\Cases\RS19-ACU-QAC.txt','r');

text = fileread('O:\Administration\BackupFiles\LogFiles\Cases\RS19-ACU-QAC.txt');

clc
fid = fopen(filename);
tline = fgetl(fid);
while ischar(tline)
    disp(tline)
    if contains(tline,'Shipment') == 1 && contains(tline,'Termination') == 1
        disp('Shipped')
    else
        disp('Not shipped')
    end
    tline = fgetl(fid);
    
    
end

fclose(fid);