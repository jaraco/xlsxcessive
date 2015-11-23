from setuptools import setup

setup(
    name="XlsXcessive",
    version="0.2.0",
    description="A Python library for writing .xlsx files.",
    packages=['xlsxcessive'],
    author="Christian Wyglendowski",
    author_email="christian@dowski.com",
    maintainer="Jason R. Coombs",
    maintainer_email="jaraco@jaraco.com",
    url="https://bitbucket.org/jaraco/xlsxcessive",
    install_requires = [
        'openpack',
        'six',
    ],
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 2.7",
        "Programming Language :: Python :: 3",
    ],
)
