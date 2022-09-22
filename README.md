# travel-statement-generator 
- this program generates random travel times and calculates corresponding values and fees into a preformatted table
 
# first-time install
- install python (don't forget to tick 'Add Python to the System Path')
```https://www.python.org/ftp/python/3.9.1/python-3.9.1-amd64.exe```
- go to the project directory
- install virtual environment
```py -m venv venv```
- activate virtual environment
```venv\Scripts\activate.bat```
- install requirements
```pip install -r requirements.txt```

# run
- go to the project directory
```py generator.py [year]```
- [year] must be a decimal integer in range 1970 and 2100

- example:
```py generator.py 2021``` (generates table for year 2021)

# notes
- the project directory MUST include ```db.xlsx``` and ```empty.xlsx``` files. Do NOT modify them!
- the program generates ```output.xlsx``` file as result