from distutils.core import setup

setup(
    name = 'TableFactory',
    version = '0.1.2',
    py_modules=['TableFactory'],
    description = 'Easily create HTML, spreadsheet, or PDF tables from common Python data sources',
    author='Kirk Strauser',
    author_email='kirk@strauser.com',
    url='http://kstrauser.github.com/tablefactory/',
    long_description='TableFactory is a simple API for creating tables in popular formats. It acts as a wrapper around other popular Python report generators and handles all the tedious, boilerplate problems of extracting columns from input data, creating the layout, applying formatting to cells, etc.',
    keywords=['reports', 'pdf', 'spreadsheet'],
    classifiers=[
        'Development Status :: 4 - Beta',
        'Environment :: Other Environment',
        'Environment :: Web Environment',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Topic :: Office/Business',
        'Topic :: Office/Business :: Financial :: Spreadsheet',
        'Topic :: Software Development :: Libraries :: Python Modules',
        'Topic :: Text Processing :: Markup :: HTML',
        ],
    install_requires=['xlwt', 'ReportLab'],
        )
