#! /usr/bin/env python
import os
from setuptools import setup

import admin_import
setup(
    name='django-admin-import',
    version = admin_import.__version__,
    description = 'Import tool attachable to almost every Django admin',
    long_description=open(os.path.join(os.path.dirname(__file__), 'README.rst')).read(),
    author = 'Marc Egli',
    author_email = 'egli@allink.ch',
    url = 'http://github.com/allink/django-admin-import/',
    license='BSD License',
    platforms=['OS Independent'],
    packages=[
        'admin_import',
    ],
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Environment :: Web Environment',
        'Framework :: Django',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: BSD License',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Topic :: Internet :: WWW/HTTP :: Dynamic Content',
    ],
    requires=[
        'Django(>=1.3)',
    ],
    include_package_data=True,
)
