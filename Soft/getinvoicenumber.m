function [invoicenumber] = getinvoicenumber(values)

invoicenumber =  num2str(values.NextInvoiceNumber);
if values.NextInvoiceNumber < 10000
    invoicenumber = ['0' num2str(values.NextInvoiceNumber)];
end
if values.NextInvoiceNumber < 1000
    invoicenumber = ['00' num2str(values.NextInvoiceNumber)];
end
if values.NextInvoiceNumber < 100
    invoicenumber = ['000' num2str(values.NextInvoiceNumber)];
end
if values.NextInvoiceNumber < 10
    invoicenumber = ['0000' num2str(values.NextInvoiceNumber)];
end

end
