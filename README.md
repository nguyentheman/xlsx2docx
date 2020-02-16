# xlsx2docx
A python3-based utility script to generate verilog design of Configuration and Status Register (CSR) from an excel input file

## Installation

This script require to install "pandas", "xlrd" and "pydocx" package to execute. They can be installed via "pip"

```
pip install pandas
pip install xlrd
pip install pydocx
```

## Usage

Using command below to generate the verilog design 

```
xlsx2rtl.py -i <excel input file> -w <word templete file> -o <output directory> -l <heading level>
```

## Examples

The example of "excel input file" and "verilog templete files" are availabled on [test](https://github.com/nguyentheman/xlsx2rtl/tree/master/test). User can execute the below command for test.

``` 
xlsx2rtl.py -i ./test/test.xlsx -w ./test/report_templete.docx -o ./test/ -l 1
``` 

The excel input file:

![Register input file format](https://github.com/nguyentheman/xlsx2docx/blob/master/docs/register_define.jpg)

The uart_csr_ug.docx output:

![Register input file format](https://github.com/nguyentheman/xlsx2docx/blob/master/docs/docx_output.jpg)

## License

This project is licensed under the MIT License - see the [LICENSE.md](https://github.com/nguyentheman/xlsx2docx/blob/master/LICENSE) file for details

