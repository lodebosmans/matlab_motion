function [invoice] = addtospecificinvoice(cc,invoice,y)
    invoice.clientspecific.(cc).counter = invoice.clientspecific.(cc).counter + 1;
    invoice.clientspecific.(cc).overview(invoice.clientspecific.(cc).counter,1:12) = invoice.filtered(y,1:12);
end
