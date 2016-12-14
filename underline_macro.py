#!/usr/bin/python3

import uno
from pprint import pprint


def log(data):
    f = open('/home/nate/git/underline-macro/log.txt', 'a')
    f.write(data + '\n')
    f.close()


def getDesktop():
    log('getDesktop')
    # If being run from inside LibreOffice, XSCRIPTCONTEXT will be defined
    if 'XSCRIPTCONTEXT' in globals():
        log("using XSCRIPTCONTEXT")
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
    log('using the socket method')
    return smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def Underline_Words():
    log('main')
    desktop = getDesktop()
    model = desktop.getCurrentComponent()

    replaceDescriptor = model.createReplaceDescriptor()

    props = structify({
                # 'CharFontName': 'Arial',
                'CharUnderline': 1
            })

    replaceDescriptor.setReplaceAttributes(props)
    replaceDescriptor.SearchRegularExpression = True
    replaceDescriptor.SearchString = '%x(.*?)x%'
    replaceDescriptor.ReplaceString = '$1'
    numReplaced = model.replaceAll(replaceDescriptor)

    print(numReplaced)
    pprint(dir(replaceDescriptor))
    print(type(replaceDescriptor.getReplaceAttributes()))


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
