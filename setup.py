"""
Setup for Python API Wrapper for MailStore Server Administration API.

This is a copy of the example Python API wrapper from MailStore, with this
setup script.

See http://en.help.mailstore.com/index.php?title=Python_API_Wrapper_Tutorial
"""

from setuptools import find_packages, setup

setup_params = dict(
    name="mailstore",
    description="Python API wrapper for the MailStore Server Administration API",
    packages=find_packages(),
    version = "0.1.0",
    url="https://github.com/renshawbay/python-mailstore-api-wrapper",
    classifiers=["License :: OSI Approved :: MIT License"],
)

if __name__ == '__main__':
    setup(**setup_params)
