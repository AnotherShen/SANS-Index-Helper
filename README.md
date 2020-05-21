# SANS Index Helper
A small tool that will take an excel template used for SANS indexing and collate up to three terms per reference into an alphabetical index using distinct text.

## Requirements
This script requires 'openpyxl'

`pip install -r requirements.txt`

## Usage
1. Create template.xlsx using the headers seen below and fill it in
2. Run sindex.py
```
$ python sindex.py -h
sindex.py -i <inputfile> [-o <outputfile>]
  
$ python sindex.py -i template
Success! -> template.xlsx processed into index_out.xlsx

$ python sindex.py -i tempalte -o namedout
Success! -> template.xlsx processed into namedout.xlsx
```
Note that .xlsx can be included at the end of each file without changing the result.

## Quirks
1. Terms are case-sensitive on purpose
2. Not every term has to be populated, blank terms will be discarded
3. Terms with an empty book and page reference will be discarded
4. References are ordered by occurence in the template

## Template Format
```
| Book | Page | Term 1 | Term 2 | Term 3 |
|  1   |  20  | $HOME  |  dig   |        |
|  1   |  52  | whoami |  dig   | $HOME  |
```
![Template XLSX](https://i.imgur.com/L2T9QTF.png "Template XLSX")

## Example Output
```
|  Term  |  Reference |
| $HOME  | 1:20, 1:52 |  
| dig    | 1:20, 1:52 |
| whoami | 1:52       |
```
![Output XLSX](https://i.imgur.com/oNrY05T.png "Output XLSX")
