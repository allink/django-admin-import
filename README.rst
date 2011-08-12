Admin Import
============

Installation
------------

**Notice**: This application is still under heavy development and every thing
can be subject of change.

1. Add admin_import to INSTALLED_APPS in your settings file.

2. Add an import to your admin like so:
 
    try:
        from admin_import.options import add_import
    except ImportError:
        pass
    else:
        add_import(InviteeAdmin)    

3. Add a button in your admin by overriding the list_view template.