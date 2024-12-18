# Farnell Email list

def checkEmailAddresses(allOrdersInCSVfile,currentEmailAddresses):
    invoiceNames = extractInvoiceNames(allOrdersInCSVfile)
    namesWithAddresses, namesWithoutAddresses = checkIfAddressesExist(invoiceNames,currentEmailAddresses)
    if len(namesWithoutAddresses) > 0:
        print('The following individuals lack Email addresses, please add them to the Sheet (2nd tab):')
        print(namesWithoutAddresses)
        return False, 'Some individuals lack Email addresses.'
    else:
        EmailAddresses = []
        for name in namesWithAddresses:
            for address in currentEmailAddresses:
                if name == address[0]:
                    EmailAddresses.append(address)
                else:
                    pass
        return True, EmailAddresses
    

def checkIfAddressesExist(invoiceNames,currentEmailAddresses):
    namesWithAddresses = []
    namesWithoutAddresses = []
    for name in invoiceNames:
        addressExists = any(name in address for address in currentEmailAddresses)
        if addressExists:
            namesWithAddresses.append(name)
        else:
            namesWithoutAddresses.append(name)
    return namesWithAddresses, namesWithoutAddresses


def extractInvoiceNames(allOrdersInCSVfile):
    invoiceNames = []
    for individual in allOrdersInCSVfile:
        name = individual[0][0] + ' ' + individual[0][1]
        invoiceNames.append(name)
    return invoiceNames


def main():
    print('This is the Email list script.')


if __name__ == '__main__':
    main()