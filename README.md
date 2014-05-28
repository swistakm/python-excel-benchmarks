# Simple benchmark for various Excel/XLSX python libraries

There is a plenty of Excel/XLSX python packages on PyPI. To name a few:
* openpyxl
* xlwt
* xlsxcessive
* xlsxwriter
* pyexcelerate
* excellent

Mostly buggy, poorly documented or not maintaned anymore. These which are
maintained seem to completely ignore problem of efficiency.

This simple benchmark tests performance of choosen python packages on
simple task of creating spreadsheet that contains 1000 x 100 cells with
different values. Nothing more.

## Run benchmark
To run benchmark simply clone repo and install requirements:

```
git clone https://github.com/swistakm/python-excel-benchmarks.git
cd python-excel-benchmarks
pip install -r requirements.txt
python benchmark.py
```

If you want to benchmark `xlsxcessive` install it manually because it is
packaged without any respect to anyone:

```
pip install xlsxcessive --allow-external xlsxcessive --allow-unverified xlsxcessive
```


## Method
Method is simple - create one spreadsheet with one sheet that contains
1000 rows and 100 columns with set of cycled values (row by row). It's a
100k cells. I know it seems many but `csv` deals with it in matter of hundredths
of a second and some folks sent a car to Mars. Expecting that export to XLSX
of 100k cells of data will take less than a second is really to much?

It seems so...

You can change parameters of benchmark. Run `python benchmark.py -h` to see
whole usage.

It will output benchmark templates to local file so you can check if output
looks as it should. I know that disk I/O will probably affect benchmark results
but I have no time or desire to dig into code of all those libs to see if they
can undestand anything more than "file name" when saving.

Each benchmark uses spreadsheet creation method that seems most
idiomatic/efficient for API of tested package. It could not be the most
efficient because most of packages lacks any sensible documentation.

Didn't tested spreadsheet creation with styling because I'm not satissfied
even with plain data.

## Findings
I run benchmarks on my machine:
```
OS: Ubuntu 14.04 trusty
Kernel: x86_64 Linux 3.12.0-5-generic
CPU: Intel Core2 Duo CPU E8200 @ 2.667GHz
RAM: 7723MB
```

Here are the results:
```
benchmark_csv                  0.026495
benchmark_excellent            1.784107
# SKIP benchmark_ooxml (ooxml is FUBAR)
benchmark_openpyxl             2.897072
benchmark_openpyxl_rows        7.025895
benchmark_pyexcelerate         0.550225
benchmark_xlsxcessive          1.430242
benchmark_xlsxwriter           1.611668
benchmark_xlwt                 1.275144
```

At the moment the best is `pyexcelerate` which is advertised as built with
speed in mind. But beware. Some of its code looks like mess (even mixed
indentation between modules) and its ["insane optimization claim"](https://github.com/kz26/PyExcelerate/blob/0.6.1/pyexcelerate/Range.py#L7)
looks like neat excuse for that code mess.
