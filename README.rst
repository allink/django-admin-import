Admin Import
============

Installation
------------

**Notice**: This application is still under heavy development and everything
is subject to change.

1. Add ``'admin_import'`` to ``INSTALLED_APPS`` in your settings file.

2. Add an import to your admin like so::

    try:
        from admin_import.options import add_import
    except ImportError:
        pass
    else:
        add_import(InviteeAdmin)

3. Add a button in your admin by overriding the ``change_list.html`` template
   for your specific model or use ``add_import(InviteeAdmin, add_button=True)``

4. If you need an other modelform than used in the admin you can override
   the ``get_admin_import_model_form` function. But it should be capable of
   excluding some files.
