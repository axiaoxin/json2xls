#!/usr/bin/env python
# -*- coding: utf-8 -*-


try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup


readme = open('README.md').read()
history = open('HISTORY.rst').read().replace('.. :changelog:', '')

requirements = [
    # TODO: put package requirements here
    'requests',
    'click',
    'xlwt',
]

test_requirements = [
    # TODO: put package test requirements here
]

setup(
    name='json2xls',
    version='0.1.0',
    description='generate excel by json',
    long_description=readme + '\n\n' + history,
    author='axiaoxin',
    author_email='254606826@qq.com',
    url='https://github.com/axiaoxin/json2xls',
    packages=[
        'json2xls',
    ],
    package_dir={'json2xls':
                 'json2xls'},
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
        "Programming Language :: Python :: 2",
        'Programming Language :: Python :: 2.6',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
    ],
    test_suite='tests',
    tests_require=test_requirements
)
