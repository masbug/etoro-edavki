from distutils.core import setup

with open("README.md", "r") as fh:
    long_description = fh.read()

setup(
    name="etoro_edavki",
    version="1.0.0",
    py_modules=["etoro_edavki"],
    python_requires=">=3",
    entry_points={
        "console_scripts": ["etoro_edavki=etoro_edavki:main", "etoro-edavki=etoro_edavki:main"]
    },
    author="Mitja Spes",
)
