#!/usr/bin/env python3
"""
Offitrans - Office文件翻译工具
"""

from setuptools import setup, find_packages
import os

# 读取README文件
current_directory = os.path.abspath(os.path.dirname(__file__))
with open(os.path.join(current_directory, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()

# 读取requirements文件
with open(os.path.join(current_directory, 'requirements.txt'), encoding='utf-8') as f:
    requirements = [line.strip() for line in f if line.strip() and not line.startswith('#')]

setup(
    name='offitrans',
    version='1.0.0',
    author='Offitrans Contributors',
    author_email='offitrans@example.com',
    description='一个强大的Office文件翻译工具库，支持PDF、Excel、PPT和Word文档的批量翻译',
    long_description=long_description,
    long_description_content_type='text/markdown',
    url='https://github.com/minglu6/Offitrans',
    packages=find_packages(),
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Topic :: Office/Business :: Office Suites',
        'Topic :: Text Processing :: Linguistic',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ],
    python_requires='>=3.7',
    install_requires=requirements,
    keywords='translation, office, excel, word, pdf, ppt, powerpoint, translate',
    project_urls={
        'Bug Reports': 'https://github.com/your-username/Offitrans/issues',
        'Source': 'https://github.com/your-username/Offitrans',
        'Documentation': 'https://github.com/your-username/Offitrans#readme',
    },
    entry_points={
        'console_scripts': [
            'offitrans=offitrans.cli:main',
        ],
    },
    include_package_data=True,
    zip_safe=False,
)