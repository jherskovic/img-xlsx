### img-xlsx.py
A silly and not very useful script to turn image files into Excel spreadsheets.
Works on macOS and Windows, and probably Linux, too. Written and tested with Python 3.7, but 3.anything recent should work.

(C) 2019 Jorge Herskovic, released into the public domain.

## Installation
Create a virtual environment like a civilized person. In Python 3 parlance, clone the project, then `cd` into it, then:
```
$ python3 -m venv virtual_environment_name
$ virtual_environment_name/bin/activate
(virtual_environment_name) $ pip install -r requirements.txt
```

## Usage
To run the script:
```
(virtual_environment_name) $ python img-xlsx.py image_filename excel_filename
```

By default, it will make a spreadsheet with 64 cells on the image's longest side. Use the `--size` parameter to change this. 256 is probably the highest reasonable number.

If you wish to limit the number of colors in the spreadsheet (perhaps for cross-stitching purposes?) use `--quantize` and a number between 2 and 255. You could use 1, I guess, but that will make a solid blob.
