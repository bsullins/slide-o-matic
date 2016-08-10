from setuptools import setup

setup(
    name="slide-o-matic",
    version="0.1",
    py_modules=['pptx'],
    author="Ben Sullins",
    author_email="ben@bensullins.com",
    description="Generate slides for Pluralsight courses",
    url="https://github.com/bsullins/slide-o-matic",
    license="MIT",
    install_requires=[
        'python-pptx',
    ],
    zip_safe=False
)

