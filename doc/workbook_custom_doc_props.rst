Workbook Custom Document Properties
===================================


It is possible to add one or more CustomDocumentProperty objects to a workbook.
These require a name (string) and can be one of 5 value types or a link.
The 5 value types are as follows:
 - bool (Boolean)
 - int (Integer)
 - float (Float/Decimal)
 - str (String)
 - datetime.datetime (Date)
Your value should be of this type or be convertible to it (such as date from iso string)

The other alternative is to link your CustomDocumentProperty to a defined named range.
If you do this, you cannot create the CustomDocumentProperty with a value.

These properties are globally for a workbook and accessed from the `custom_doc_props` attribute.


Sample use
----------

Looping over all CustomDocumentProperties ("custom_doc_props")::

    for prop in wb.custom_doc_props:
        print(f"{prop.name}: {prop.value}")

Adding a new CustomDocumentProperty::

    # if you only specify the first two arguments, we take the type of the value as is
    wb.custom_doc_props.add('TestDocProp2', "foo"))

    # if you specify the third argument, you can convert a type, if it is convertible
    # we also allow a string to represent the types as follows
    # ("int", "float", "str", "bool", "date"), or you can pass in the actual types,
    # such as (int, float, str, bool, datetime.datetime)
    wb.custom_doc_props.add('TestDocProp3', "2020-08-24T20:19:22Z", "date")

    # in total there are 4 possible arguments:
    wb.custom_doc_props.add(PropName, PropVal=None, PropType=None, LinkTarget=None)




Deleting all existing CustomDocumentProperties and adding new ones
------------------------------------------------------------------

.. testcode::

    # delete any existing CustomDocumentProperties
    for prop_name in wb.custom_doc_props.namelist():
        del wb.custom_doc_props[prop_name]

    # add one of each type
    wb.custom_doc_props.add("PropName1", datetime.datetime(2020, 8, 24, hour=20, minute=19, second=22))
    wb.custom_doc_props.add("PropName2", 2.5)
    wb.custom_doc_props.add("PropName3", 2)
    wb.custom_doc_props.add("PropName4", "Foo")
    wb.custom_doc_props.add("PropName5", True)

    # add a string, which you want to convert to a float and save:
    wb.custom_doc_props.add("PropName6", "2.5", float)
    # check the converted value
    prop = wb.custom_doc_props["PropName6"]
    print(f"{prop.name}: {prop.value} {type(prop.value)}")

    # add a CustomDocumentProperty with a link instead of a value
    wb.custom_doc_props.add("PropName7", LinkTarget="ExampleName")

    # save the file
    wb.save('outfile.xlsx')


.. testoutput::

    PropName6: 2.5 <class 'float'>
