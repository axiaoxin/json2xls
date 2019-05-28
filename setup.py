#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import re

try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

readme = 'https://github.com/axiaoxin/json2xls/blob/master/README.md'

with open(os.path.join(os.path.dirname(__file__),
                       'json2xls/__init__.py')) as f:
    version = re.search(r'__version__ = \'(.*?)\'', f.read()).group(1)

requirements = ["click", "requests", "xlwt"]

test_requirements = [
    # TODO: put package test requirements here
]

setup(
    name='json2xls',
    version=version,
    description='generate excel by json',
    long_description=readme,
    author='axiaoxin',
    author_email='254606826@qq.com',
    url='https://github.com/axiaoxin/json2xls',
    packages=[
        'json2xls',
    ],
    package_dir={'json2xls': 'json2xls'},
    include_package_data=True,
    install_requires=requirements,
    license="BSD",
    zip_safe=False,
    keywords='json2xls',
    classifiers=[
        'Development Status :: 2 - Pre-Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: BSD License',
        'Natural Language :: English',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3.6',
    ],
    test_suite='tests',
    tests_require=test_requirements,
    entry_points={'console_scripts': ['json2xls = json2xls.json2xls:make']})
