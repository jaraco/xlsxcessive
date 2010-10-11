from setuptools import setup

setup(
    name="XlsXcessive",
    version="0.1.5",
    description="A Python library for writing .xlsx files.",
    packages=['xlsxcessive'],
    author="Christian Wyglendowski",
    author_email="christian@dowski.com",
    url="http://xlsx.dowski.com/",
    download_url="http://bitbucket.org/dowski/xlsxcessive/downloads",
    license="MIT",
    install_requires = [
        'openpack',
    ],
)
