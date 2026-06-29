from setuptools import setup

setup(
    name='otds-bulk-tool',
    version='1.0.0',
    description='OpenText OTDS 25.x Bulk Users & Groups Provisioning Tool',
    author='Naveen K. Manam',
    python_requires='>=3.8',
    py_modules=['otds_bulk_tool'],
    install_requires=[
        'requests>=2.28.0',
        'pandas>=1.5.0',
        'openpyxl>=3.0.10',
    ],
    entry_points={
        'console_scripts': [
            'otds-bulk-tool=otds_bulk_tool:main',
        ],
    },
    classifiers=[
        'Programming Language :: Python :: 3',
        'Operating System :: OS Independent',
        'Topic :: System :: Systems Administration',
    ],
)
