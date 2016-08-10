# slide-o-matic
Generates powerpoint slides from a json file

# setup
- Have the latest version of Python installed

## mac users
- Install dependencies via `python setup.py install --user`
- You'll need Python installed and then run setup.py
- You may also need to install the XCode command line tools if you've not done so already

# dependencies
- [python-pptx](https://pypi.python.org/pypi/python-pptx)

# usage
1. Open the Excel file and enter your module info (layouts: 0=title, 2=section header)
2. Update profile.png to your image
3. Copy the `sample-config.py` file to `config.py` and set the appropriate values
4. Copy/paste the JSON generated in the Excel file to `gen-slides.py`
5. Run `python gen-slides.py`
6. Revel in your accomplishments
7. Fork this repo to improve this solution so we can all save time typing into slides
