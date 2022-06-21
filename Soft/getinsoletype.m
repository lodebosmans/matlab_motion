function [A2] = getinsoletype(insoletype,basetype)

% Define the second part of the AA part of the itemnumber

% base type needs to be included as well !!!
if strcmp(basetype,'Normal') == 1 || strcmp(basetype,'Ortho') == 1 || strcmp(basetype,'-') == 1
    if strcmp('Running',insoletype) == 1 || strcmp('Comfort',insoletype) == 1 || strcmp('Narrow',insoletype) == 1 || strcmp('Soccer',insoletype) == 1 || strcmp('Golf',insoletype) == 1 || strcmp('Wide',insoletype) == 1
        A2 = '0';
    elseif strcmp('Alpine ski',insoletype) == 1 || strcmp('Nordic ski',insoletype) == 1 || strcmp('Cycling',insoletype) == 1
        A2 = '1';
    else
        disp(['Insole type ' insoletype ' not recognized']);
        clear all
    end
elseif strcmp(basetype,'Normalslim') == 1 || strcmp(basetype,'Orthoslim') == 1
    if strcmp('Running',insoletype) == 1 || strcmp('Comfort',insoletype) == 1 || strcmp('Narrow',insoletype) == 1 || strcmp('Soccer',insoletype) == 1 || strcmp('Golf',insoletype) == 1 || strcmp('Wide',insoletype) == 1
        A2 = '2';
    elseif strcmp('Alpine ski',insoletype) == 1 || strcmp('Nordic ski',insoletype) == 1 || strcmp('Cycling',insoletype) == 1
        A2 = '3';
    else
        disp(['Insole type ' insoletype ' not recognized']);
        clear all
    end
else
    disp(['Base type ' basetype ' not recognized']);
    clear all
end