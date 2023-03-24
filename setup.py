import os
from codecs import open
from setuptools import setup, find_packages


root_dir = os.path.abspath(os.path.dirname(__file__))
# def _requirements():
#     requirements_file_path = os.path.join(root_dir, 'requirements.txt')
#     return [name.rstrip() for name in open(requirements_file_path).readlines()]
#     install_requires = _requirements()

# Get the long description from the README file
readme_file_path = os.path.join(root_dir, 'README.rst')
with open(readme_file_path, encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='mael',
    packages=find_packages(),
    # package_dir={'': 'mael'},
    # packages=find_packages(where='mael'),
    version='0.0.3.30',
    py_modules=['mael', 'mael.main', 'mael.initializer', 'mael.excel_builder', 'mael.column_config'],
    install_requires=[
        'pyyaml',
        'openpyxl'
    ],
    entry_points={
        'console_scripts': ['mael=mael.main:main'],
    },
    license='MIT',
    author='Kenji Otsuka',
    author_email='kok.fdcm@gmail.com',
    description='A tool to convert markdown file to excel.',
    long_description=long_description,
    url='https://github.com/KenjiOhtsuka/mael',
    include_package_data=True,
    # entry_points='''
    #     [console_scripts]
    #     myproject=myproject:cli
    # ''',
)
