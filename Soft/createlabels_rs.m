function createlabels_rs()

load Temp\values.mat values

firstletters = ['R','F'];
for fl = 1:size(firstletters,2)
    cfl = firstletters(1,fl);
    % Create text file for printing label code
    fid = fopen(['Output\Labels\results_labels_' cfl 'S.txt'],'wt');
    fprintf(fid, '${');
    fprintf(fid, '\n');
    disp(['Creating ' cfl 'S-labels - please wait']);
    % Create for loop to iterate over all IDs
    for rsindex = values.InitialRSlabelnr:values.FinalRSlabelnr
        % Generate the final number
        if rsindex < 10
            finalnr = ['0000' num2str(rsindex)];
        elseif rsindex < 100
            finalnr = ['000' num2str(rsindex)];
        elseif rsindex < 1000
            finalnr = ['00' num2str(rsindex)];
        elseif rsindex < 10000
            finalnr = ['0' num2str(rsindex)];
        else
            clear all
        end
        % Generate the entire rscode
        rscode = [cfl 'S-' values.RequestedYear(3:4) values.RequestedMonth '-' finalnr];
        % Orientation for 180° rotated (^POI) print
        code.part1 = ['^XA^CFU^POI^FO80,100,0^FD' rscode '^FS^FWZ,0^FO510,20^BQN,2,10^FDQA,'  rscode '^FS^XB^XZ'];
        code.part2 = ['^XA^CFU^POI^FO80,100,0^FD' rscode '^FS^FWZ,0^FO510,20^BQN,2,10^FDQA,'  rscode '^FS^XZ'];
        if rsindex < values.FinalRSlabelnr
            fprintf(fid, code.part1);
            fprintf(fid, '\n');
            fprintf(fid, code.part1);
            fprintf(fid, '\n');
        end
        if rsindex == values.FinalRSlabelnr
            fprintf(fid, code.part1);
            fprintf(fid, '\n');
            fprintf(fid, code.part2);
            fprintf(fid, '\n');
        end
        clear code
    end
    fprintf(fid, '}$');
    fclose(fid);
end

end