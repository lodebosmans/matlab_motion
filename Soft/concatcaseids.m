function [output] = concatcaseids(input)

input = input';
input(2,:) = {'(1),'} ;
output = [input{:}];

end