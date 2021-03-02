from distutils.core import setup
import setuptools

long_description = "eToro -> eDavki konverter"

setup(
    name="etoro_edavki",
    version="1.0.0",
    py_modules=["etoro_edavki"],
    python_requires=">=3",
    entry_points={
        "console_scripts": ["etoro_edavki=etoro_edavki:main", "etoro-edavki=etoro_edavki:main"]
    },
    author="Mitja Spes",
    install_requires=[
	    'openpyxl>=3.0.5',
		'openpyxl-templates>=0.2.5',
		'prettytable>=2.0.0',
		'future>=0.18.2'
	]
)
