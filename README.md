# Excel to digital calendar converter
At my music band our appointment schedule gets distributed in a pretty cursed
excel file layout. I wrote this script to convert our appointment schedule
into an `.ics` file which can be imported into e.g. outlook calendars.

*This is just a quick evening project to automate some boring stuff. This code 
is not pretty at all and I'm aware of it.*

## Running
Put the appointment schedule file into the git projects root directory and
rename it to `input.xlsx`. Then run the command below. A file named
`output.ics` will be created by the python script. You can import this e.g. into
your outlook calendar.

`nix develop --command python excelToIcs.py`

