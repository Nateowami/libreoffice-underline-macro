import uno
from pprint import pprint

# get the uno component context from the PyUNO runtime
localContext = uno.getComponentContext()

# create the UnoUrlResolver
resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext)

# connect to the running office
ctx = resolver.resolve(
        "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
smgr = ctx.ServiceManager

# get the central desktop object
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

# desktop = XSCRIPTCONTEXT.getDesktop()
model = desktop.getCurrentComponent()


def main():

    # text = model.Text

    replaceDescriptor = model.createReplaceDescriptor()

    # print(dir(sd.ReplaceAttributes))

    # struct = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
    # struct.Name = 'CharFontName'
    # struct.Value = 'Arial'
    # props = tuple([struct])
    props = structify({'CharFontName': 'Arial', 'CharFontUnderline': 1})

    replaceDescriptor.setReplaceAttributes(props)
    replaceDescriptor.SearchRegularExpression = True
    replaceDescriptor.SearchString = 'H(.{3})o'
    replaceDescriptor.ReplaceString = 'Do you like j$1o'
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


main()
