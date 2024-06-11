from setuptools import setup, find_packages

setup(
    name='tfstate_to_excel',
    version='0.1.0',
    description='Extract Terraform state to an Excel workbook',
    author='Alex KM',
    author_email='me@alexkm.com',
    packages=find_packages(),
    install_requires=[
        'openpyxl',
    ],
    entry_points={
        'console_scripts': [
            'tfstate_to_excel=tfstate_excel_extractor:main',
        ],
    },
)

