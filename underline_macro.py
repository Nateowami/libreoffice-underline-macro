#!/usr/bin/python3

import uno


# Logging util for debugging when run from within LibreOffice.
def log(data):
    # Logging only works on *nix systems
    f = open('~/log.txt', 'a')
    f.write(data + '\n')
    f.close()


def getDesktop():
    # If being run from inside LibreOffice, XSCRIPTCONTEXT will be defined
    if 'XSCRIPTCONTEXT' in globals():
        return XSCRIPTCONTEXT.getDesktop()  # NOQA to disable Flake8 here

    # Otherwise, if we're running form the command line, we have to connect to
    # a running instance of LibreOffice. Libreoffice must be started to listen
    # on a socket that we connect to in order for this to work.

    localContext = uno.getComponentContext()
    # create the UnoUrlResolver
    resolver = localContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext)
    # connect to the running office
    ctx = resolver.resolve(
            "uno:socket,host=localhost,port=2002;urp;"
            "StarOffice.ComponentContext")
    smgr = ctx.ServiceManager
    # get the central desktop object
    return smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def Underline_Words():
    desktop = getDesktop()
    model = desktop.getCurrentComponent()

    # Remove closing "tags" that are immediately followed by opening tags.
    # Otherwise there could be tiny gaps in the underlining.
    apply_regex(model, 'x% ?%x', '')
    apply_regex(model, '%x(.*?)x%', '$1', {'CharUnderline': 1})


def apply_regex(model, search, replace, replaceAttrs={}):
    replaceDescriptor = model.createReplaceDescriptor()

    props = structify(replaceAttrs)
    replaceDescriptor.setReplaceAttributes(props)

    replaceDescriptor.SearchRegularExpression = True
    replaceDescriptor.SearchString = search
    replaceDescriptor.ReplaceString = replace

    numReplaced = model.replaceAll(replaceDescriptor)
    return numReplaced


def structify(keypairs):
    result = []
    for key, value in keypairs.items():
        struct = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
        struct.Name = key
        struct.Value = value
        result.append(struct)
    return tuple(result)


g_exportedScripts = (Underline_Words, )

if __name__ == "__main__":
    Underline_Words()
