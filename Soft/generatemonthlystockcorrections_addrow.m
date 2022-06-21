function [output] = generatemonthlystockcorrections_addrow(output,info,values,itemnumber,correction,amount)

output_row = size(output,1) + 1;

% Fill in the fixed stuff
output(output_row,1) = {info.counter};
output(output_row,2) = info.caseid;
output(output_row,3) = info.left;
output(output_row,4) = info.right;
output(output_row,5) = {info.builtactor};
output(output_row,6) = {info.itemnumber_full};

output(output_row,8) = {info.output_date};
output(output_row,9) = {correction};
output(output_row,10) = {['20' values.y values.mo values.d '_StockCorLow']}; % Use day of creation of file
output(output_row,11) = {itemnumber};
output(output_row,12) = {''};
output(output_row,13) = {'BEPROD'};
output(output_row,14) = {'BEL'};
output(output_row,15) = {amount};

end