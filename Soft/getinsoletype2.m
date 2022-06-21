function [type2] = getinsoletype2(itemnr,material,assembly,basetype)

if strcmp(itemnr(3:3),'0') == 1 || strcmp(itemnr(3:3),'2') == 1
    part2 = 'Semi';    
elseif strcmp(itemnr(3:3),'1') == 1 || strcmp(itemnr(3:3),'3') == 1
    part2 = 'Full';    
else
    disp('Fatal error');
    clear all
end

if strcmp(basetype,'Normal') == 1 || strcmp(basetype,'Ortho') == 1 || strcmp(basetype,'-') == 1
    part1 = 'Reg';
elseif strcmp(basetype,'Normalslim') == 1 || strcmp(basetype,'Orthoslim') == 1
    part1 = 'Slim';
else
    disp('Fatal error');
    error('The insole type could not be recognized!');
end

typeshort = [part1 part2];

% if no toplayer
if strcmp(assembly,'No top layer') == 1 || strcmp(itemnr(4:5),'00') == 1
    type2 = ['TL_' typeshort '_None'];
    % If non-assembled EVA
elseif strcmp(assembly,'Not assembled - full length') == 1 && ...
        (strcmp(material,'EVA') == 1 || strcmp(material,'EVA Synthetic Leather') == 1)
    type2 = ['TL_' typeshort '_Eva_NotAs'];
    % If assembled EVA
elseif (strcmp(assembly,'Assembled - full length') == 1 || strcmp(assembly,'Assembled - 3/4th length') == 1) && ...
        (strcmp(material,'EVA') == 1 || strcmp(material,'EVA Synthetic Leather') == 1)
    type2 = ['TL_' typeshort '_Eva_As'];
    % If non-assembled PU soft
elseif strcmp(assembly,'Not assembled - full length') == 1 && ...
        (strcmp(material,'PU Soft') == 1 || strcmp(material,'PU Soft Synthetic Leather') == 1)
    type2 = ['TL_' typeshort '_Pusoft_NotAs'];
    % If assembled PU soft
elseif (strcmp(assembly,'Assembled - full length') == 1 || strcmp(assembly,'Assembled - 3/4th length') == 1) && ...
        (strcmp(material,'PU Soft') == 1 || strcmp(material,'PU Soft Synthetic Leather') == 1)
    type2 = ['TL_' typeshort '_PuSoft_As'];
    % If non-assembled EVA cabon
elseif strcmp(assembly,'Not assembled - full length') == 1 ...
        && strcmp(material,'EVA Carbon') == 1
    type2 = ['TL_' typeshort '_EvaCarbon_NotAs'];
    % If assembled EVA cabon
elseif (strcmp(assembly,'Assembled - full length') == 1 || strcmp(assembly,'Assembled - 3/4th length') == 1) && ...
        strcmp(material,'EVA Carbon') == 1
    type2 = ['TL_' typeshort '_EvaCarbon_As'];
else
    disp('Combination of top layer and assembly not recognized.');
    clear all
end


end