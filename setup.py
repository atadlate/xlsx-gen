__author__ = 'mcharbit'

import ez_setup
ez_setup.use_setuptools()

from setuptools import setup, find_packages

setup(
    name='XlsxGen',
    version='0.1',

    packages=find_packages(),
    scripts=['bin/utils.py'],

    package_data = {
        # Include template files:
        '': ['Template.xlsx', 'Template_no_visible_styles.xlsx', 'Template_visible_styles.xlsx'],
    },
    author='mcharbit',
    author_email='mcharbit@pentalog.fr',
    license='GPL',
    description='Simple and fast xlsx files generator',
    long_description=open('README.txt').read(),
)

