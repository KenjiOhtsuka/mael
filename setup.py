from setuptools import setup

setup(
    name='mael',
    version='0.1',
    # py_modules=['mael'],
    install_requires=[
        'openpyxl',
    ],
    entry_points={
        'console_scripts': ['mael=main:main'],
    }
    # entry_points='''
    #     [console_scripts]
    #     myproject=myproject:cli
    # ''',
)

