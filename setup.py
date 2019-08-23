#!/usr/bin/env python
# -*- coding:utf-8 -*-
"""
    Excel Template
    ~~~~~~~~~~~~~~~~~~~
    Create Excel file according to the Excel template which fits a specific format
    :copyright: (c) 2019 by Junjie Qiu
    :license: MIT, see LICENSE for more details
"""
from os import path
from codecs import open
from setuptools import setup

basedir = path.abspath(path.dirname(__file__))

# Get the long description from the README file
with open(path.join(basedir, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='excel-template',  # 包名称
    version='0.1.5',  # 版本
    url='https://github.com/qjjayy/excel_template',
    license='MIT',
    author='Junjie Qiu',
    author_email='xiaohaixie@qq.com',
    description='Create Excel file according to the Excel template which fits a specific format',
    long_description=long_description,
    long_description_content_type='text/markdown',  # 长描述内容类型
    platforms='any',
    packages=['excel_template'],  # 包含的包列表
    zip_safe=False,
    include_package_data=False,
    install_requires=['openpyxl'],
    keywords='excel template render',
    classifiers=[
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 2.7'
    ]
)
