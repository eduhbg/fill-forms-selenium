# :robot: Fill forms with Selenium and Python

### What is Selenium?

Selenium is a free (open-source) automated testing framework used to validate web applications across different browsers and platforms.

### About this project

The project aims to study Selenium test tool and its use with Python for form filling.

I started studying Selenium after I needed to register some data in a web application of my work. I had a spreadsheet with more than 8,000 entries and I used the tool to register some of them.

There are other faster ways to fill out forms, like using `requests` in Python. However, it was interesting to see how this tool works and how could be applied in other projects.

## Installation

To execute this project you will need:

- Python ([installation instructions](https://www.python.org/downloads))
- Selenium and Driver ([installation instructions](https://selenium-python.readthedocs.io/installation.html) - [driver repository](https://github.com/mozilla/geckodriver/releases))
- Openpyxl ([installation instructions](https://openpyxl.readthedocs.io/en/stable/))

### Problem with PATH

I have one problem during the execution of my script and was because the Driver needed to be in PATH.

To solve this problem, I put the file in the Python's folder and worked fine (the location of Python in your computer is something like this: `C:\Users\YOURUSER\AppData\Local\Programs\Python\Python310`)

## Usage

In this project there is already a form created by me to perform the tests, as well as a file with data, `financial_sample.xlsx`, which will be used in the execution of this project.

Just run the script `app.py` and see the result.

## Result

![Result of the project](img/selenium.gif)
